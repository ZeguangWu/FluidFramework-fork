/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

// Measures how the Quill-on-SharedTree editor scales with document size. Mirrors ScriptorPad's `pnpm test:perf`
// sweep (office-bohemia, apps/scriptor-pad/src/automation/perfScaling/PerfScaling.perf-test.ts): for every paragraph
// count it opens a fresh page, runs the warm-up runs, discards their statistics, runs the measured runs and writes
// `results.json` (same row format as ScriptorPad) and `samples.json` (every individual result).
//
// From examples/data-objects/text-editor:
//   pnpm exec webpack --config perf/webpack.perf.cjs
//   node perf/quillPerfScaling.mjs
//
// Environment variables:
//   QUILL_PERF_SIZES     Paragraph counts. Default: 10,20,30,40,50,60,70,80,90,100
//   QUILL_PERF_RUNS      Measured runs per paragraph count. Default: 10
//   QUILL_PERF_WARMUP    Unmeasured runs before each paragraph count. Default: 1
//   QUILL_PERF_CHROMIUM  Chromium executable. Default: the one installed for the Playwright in use.
//   QUILL_PERF_PLAYWRIGHT  Path of a Playwright module entry point to use instead of this package's
//                        @playwright/test, e.g. the one ScriptorPad uses, so both run the same browser build.
//   QUILL_PERF_OUTPUT    Output directory. Default: perf/results/<timestamp>
//   QUILL_PERF_TREE      "default" to use SharedTree's default configuration (as Scriptor does), or "forest" to use the
//                        default configuration with only the example app's optimized forest. Default: the app's
//                        configuration.
//   QUILL_PERF_BACKEND   Back end name written to the results. Default: QuillSharedTree

import { createReadStream, existsSync, mkdirSync, writeFileSync } from "node:fs";
import { createServer } from "node:http";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const { chromium } = await import(
	process.env.QUILL_PERF_PLAYWRIGHT
		? pathToFileURL(process.env.QUILL_PERF_PLAYWRIGHT).href
		: "@playwright/test"
);

const perfDirectory = path.dirname(fileURLToPath(import.meta.url));
const distDirectory = path.join(perfDirectory, "dist");

function parseInteger(text, name, minimum) {
	const value = Number(text);
	if (!Number.isInteger(value) || value < minimum) {
		throw new Error(`${name} must be an integer of at least ${minimum}, but is "${text}".`);
	}
	return value;
}

function parseList(text, name, defaults) {
	if (text === undefined || text.trim() === "") {
		return defaults;
	}
	return text.split(",").map((item) => parseInteger(item.trim(), name, 1));
}

const treeConfiguration = process.env.QUILL_PERF_TREE?.trim() || undefined;
if (treeConfiguration !== undefined && treeConfiguration !== "default" && treeConfiguration !== "forest") {
	throw new Error(`QUILL_PERF_TREE must be "default", "forest" or unset, but is "${treeConfiguration}".`);
}
const backEndName = process.env.QUILL_PERF_BACKEND?.trim() || "QuillSharedTree";

const settings = {
	scriptName: "Perf-Basic",
	backEnds: [backEndName],
	treeConfiguration: treeConfiguration ?? "app",
	paragraphCounts: parseList(
		process.env.QUILL_PERF_SIZES,
		"QUILL_PERF_SIZES",
		[10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
	),
	runsPerPoint: parseInteger(process.env.QUILL_PERF_RUNS ?? "10", "QUILL_PERF_RUNS", 1),
	warmupRuns: parseInteger(process.env.QUILL_PERF_WARMUP ?? "1", "QUILL_PERF_WARMUP", 0),
};

if (!existsSync(path.join(distDirectory, "index.html"))) {
	throw new Error(
		"The harness is not built. Run `pnpm exec webpack --config perf/webpack.perf.cjs` first.",
	);
}

const outputDirectory = path.resolve(
	process.env.QUILL_PERF_OUTPUT ??
		path.join(perfDirectory, "results", new Date().toISOString().replaceAll(/[:.]/g, "-")),
);
mkdirSync(outputDirectory, { recursive: true });
console.log(`Performance scaling results: ${outputDirectory}`);

const contentTypes = {
	".html": "text/html; charset=utf-8",
	".js": "text/javascript; charset=utf-8",
	".css": "text/css; charset=utf-8",
};
const server = createServer((request, response) => {
	const { pathname } = new URL(request.url ?? "/", "http://localhost");
	const file = path.join(
		distDirectory,
		pathname === "/" ? "index.html" : decodeURIComponent(pathname.slice(1)),
	);
	if (!file.startsWith(distDirectory) || !existsSync(file)) {
		response.writeHead(404);
		response.end();
		return;
	}
	response.writeHead(200, {
		"content-type": contentTypes[path.extname(file)] ?? "application/octet-stream",
	});
	createReadStream(file).pipe(response);
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const baseUrl = `http://127.0.0.1:${server.address().port}`;

// Same flags as ScriptorPad's Playwright configuration, which enable forced GC and precise heap sizes.
const browser = await chromium.launch({
	headless: true,
	executablePath: process.env.QUILL_PERF_CHROMIUM || undefined,
	args: ["--js-flags=--expose-gc", "--enable-precise-memory-info"],
});
const browserVersion = browser.version();
console.log(`Chromium ${browserVersion}`);
let summarizerTeardownErrors = 0;

async function measurePoint(paragraphCount) {
	// A fresh page per point keeps the heap of one measurement from affecting the next.
	const context = await browser.newContext();
	try {
		const page = await context.newPage();
		page.on("pageerror", (error) => console.log(`[page error] ${error.message}`));
		page.on("console", (message) => {
			if (message.type() !== "error") {
				return;
			}
			const text = message.text();
			// Expected: the summarizer of the previous run's closed container fails its last summary because the
			// harness deletes the documents when the next run starts.
			if (text.includes("noninteractive/summarizer") && text.includes("Blob not found")) {
				summarizerTeardownErrors += 1;
				return;
			}
			console.log(`[console.error] ${text.slice(0, 2000)}`);
		});
		await page.goto(
			`${baseUrl}/?fluidClient=ephemeral${treeConfiguration === undefined ? "" : `&tree=${treeConfiguration}`}`,
		);
		await page.waitForFunction(() => window.quillPerf !== undefined);

		const run = async (repeatCount) => {
			const report = await page.evaluate((request) => window.quillPerf.run(request), {
				paragraphCount,
				repeatCount,
			});
			if (report.status !== "passed") {
				throw new Error(report.error ?? `The run ended with status ${report.status}.`);
			}
		};

		if (settings.warmupRuns > 0) {
			await run(settings.warmupRuns);
		}
		// Only the measured runs count: discard what the warm-up collected.
		await page.evaluate(() => window.quillPerf.resetPerformanceStatistics());
		await run(settings.runsPerPoint);

		const statistics = await page.evaluate(() => window.quillPerf.getPerformanceStatistics());
		const pointSamples = await page.evaluate(() => window.quillPerf.getSamples());
		return {
			rows: statistics.map((entry) => ({
				paragraphCount,
				backEnd: backEndName,
				measureName: entry.measureName,
				counter: entry.counter,
				sampleCount: entry.sampleCount,
				mean: entry.mean,
				standardDeviation: entry.standardDeviation,
				min: entry.min,
				max: entry.max,
			})),
			samples: pointSamples.map((sample) => ({ paragraphCount, ...sample })),
		};
	} finally {
		await context.close();
	}
}

const rows = [];
const samples = [];
const failures = [];
try {
	for (const paragraphCount of settings.paragraphCounts) {
		const startedAt = Date.now();
		try {
			const point = await measurePoint(paragraphCount);
			rows.push(...point.rows);
			samples.push(...point.samples);
			console.log(
				`${backEndName}, ${paragraphCount} paragraphs: ${settings.runsPerPoint} runs in ${((Date.now() - startedAt) / 1000).toFixed(1)}s`,
			);
		} catch (error) {
			const message = `${backEndName}, ${paragraphCount} paragraphs: ${error instanceof Error ? error.message : String(error)}`;
			console.log(`FAILED ${message}`);
			failures.push(message);
		}
		// Written after every point, so a long sweep that is interrupted still leaves its results.
		writeFileSync(
			path.join(outputDirectory, "results.json"),
			JSON.stringify({ settings, browserVersion, rows }, undefined, 2),
		);
		writeFileSync(
			path.join(outputDirectory, "samples.json"),
			JSON.stringify({ settings, browserVersion, samples }, undefined, 2),
		);
	}
} finally {
	await browser.close();
	server.close();
}

console.log(
	`Ignored ${summarizerTeardownErrors} expected summarizer errors from closed documents.`,
);
if (failures.length > 0) {
	console.log(`Points that could not be measured:\n${failures.join("\n")}`);
	process.exitCode = 1;
}
