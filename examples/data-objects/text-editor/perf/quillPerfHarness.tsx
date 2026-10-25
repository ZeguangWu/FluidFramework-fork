/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

/**
 * Browser-side performance harness for the Quill editor on SharedTree (`FormattedMainView` from
 * `@fluidframework/quill-react`), configured like the text-editor example app (same schema, SharedTree
 * configuration, service options and undo/redo manager).
 *
 * It reproduces ScriptorPad's `Perf-Basic` CommandScript (office-bohemia, apps/scriptor-pad) and the semantics of
 * its `PerformanceStart` / `PerformanceEnd` commands, so the results can be compared with Scriptor's Solo, Fluid and
 * SharedTree back ends:
 *
 * - Every run starts by closing the previous editor and resetting the in-memory service, then opens a new document
 * with a single newly mounted editor (ScriptorPad clears its document server before every script run).
 * - After every command the harness waits until the container has no unacknowledged ops and then yields one
 * macrotask, like ScriptorPad's `syncModelForSharedSessions`. Measurements therefore include the op round trip
 * through the local in-memory service.
 * - Execution time is `performance.now()` from `PerformanceStart` to `PerformanceEnd`. Working set is the change in
 * `performance.memory.usedJSHeapSize`, each value read right after a forced `gc()`.
 * Chromium must be launched with `--js-flags=--expose-gc --enable-precise-memory-info`.
 *
 * The page must be loaded with `?fluidClient=ephemeral` so the example service client uses the purely in-memory
 * service, like ScriptorPad's local delta connection server.
 */

import { defaultServiceOptions, getExampleServiceClient } from "@fluid-example/example-utils";
import {
	FormattedMainView,
	FormattedTextDefault,
	// eslint-disable-next-line import-x/no-internal-modules
} from "@fluidframework/quill-react/internal";
import {
	createUndoRedo,
	toPropTreeNode,
	type UndoRedo,
	// eslint-disable-next-line import-x/no-internal-modules
} from "@fluidframework/react/internal";
import { TreeViewConfiguration, type ITree } from "fluid-framework";
import {
	asAlpha,
	configuredSharedTree,
	defineDataStore,
	FluidClientVersion,
	ForestTypeOptimized,
	incrementalEncodingPolicyForAllowedTypes,
	incrementalSummaryHint,
	instantiateTreeFirstTime,
	PlainText,
	SchemaFactoryAlpha,
	sharedObjectRegistryFromIterable,
	TreeCompressionStrategy,
	TreeViewConfigurationAlpha,
	type FluidContainer,
	type TreeViewAlpha,
} from "fluid-framework/alpha";
import Quill from "quill-next";
// eslint-disable-next-line import-x/no-internal-modules, import-x/no-unassigned-import
import "quill-next/dist/quill.snow.css";
import { createRoot, type Root } from "react-dom/client";

// #region Document schema and data store, identical to src/app.tsx

const sf = new SchemaFactoryAlpha("com.fluidframework.example.text-editor");

class TextEditorRoot extends sf.objectAlpha("TextEditorRoot", {
	plainText: sf.types([PlainText.Tree], { custom: { [incrementalSummaryHint]: true } }),
	formattedText: sf.types([FormattedTextDefault.Tree], {
		custom: { [incrementalSummaryHint]: true },
	}),
}) {}

const treeConfig = new TreeViewConfiguration({ schema: TextEditorRoot });

/**
 * The example app's SharedTree configuration by default. With `?tree=default`, SharedTree's default configuration (the
 * same as the `SharedTree` export), which Scriptor's SharedTree back end uses. With `?tree=forest`, the default
 * configuration with only the app's optimized forest. The variants show what the configuration contributes.
 */
const treeOption = new URLSearchParams(location.search).get("tree");
const SharedTree =
	treeOption === "default"
		? configuredSharedTree({})
		: treeOption === "forest"
			? configuredSharedTree({ forest: ForestTypeOptimized })
			: configuredSharedTree({
				forest: ForestTypeOptimized,
				treeEncodeType: TreeCompressionStrategy.CompressedIncremental,
				shouldEncodeIncrementally: incrementalEncodingPolicyForAllowedTypes(
					new TreeViewConfigurationAlpha({ schema: TextEditorRoot }),
				),
				minVersionForCollab: FluidClientVersion.v2_74,
			});

interface TextEditorData {
	tree: ITree;
	treeView: TreeViewAlpha<typeof TextEditorRoot>;
}

const TextEditorDataStore = defineDataStore<TextEditorData, ITree>({
	type: "text-editor",
	registry: sharedObjectRegistryFromIterable([SharedTree]),
	instantiateFirstTime: async (rootCreator, creator) =>
		instantiateTreeFirstTime(rootCreator, creator, SharedTree, {
			config: treeConfig,
			initializer: () =>
				new TextEditorRoot({
					plainText: PlainText.Tree.fromString(""),
					formattedText: FormattedTextDefault.Tree.fromString(""),
				}),
		}),
	view: async (tree) => ({ tree, treeView: asAlpha(tree.viewWith(treeConfig)) }),
});

// #endregion

// #region Measurement

type PerformanceCounter = "executionTimeMs" | "workingSetBytes";

const performanceCounters: readonly PerformanceCounter[] = ["executionTimeMs", "workingSetBytes"];

/** The back end name reported for this harness, next to ScriptorPad's Solo, Fluid and SharedTree. */
const backEnd = "QuillSharedTree";

/** Statistics of one counter of one measurement, in the shape of ScriptorPad's `PerformanceCounterStatistics`. */
interface PerformanceCounterStatistics {
	readonly backEnd: string;
	readonly measureName: string;
	readonly counter: PerformanceCounter;
	readonly sampleCount: number;
	readonly mean: number;
	/** Sample standard deviation (n - 1 denominator). Zero when there is a single sample. */
	readonly standardDeviation: number;
	readonly min: number;
	readonly max: number;
}

/** One `PerformanceEnd` result. */
interface PerformanceSample {
	readonly run: number;
	readonly measureName: string;
	readonly executionTimeMs: number;
	readonly workingSetBytes: number;
	readonly initialWorkingSetBytes: number;
	readonly finalWorkingSetBytes: number;
}

/** Running mean and variance (Welford's algorithm), as in ScriptorPad's `PerformanceStatistics`. */
interface RunningStatistics {
	sampleCount: number;
	mean: number;
	sumOfSquares: number;
	min: number;
	max: number;
}

// Keyed by measure name, then counter. Maps keep first-reported order.
const statistics = new Map<string, Map<PerformanceCounter, RunningStatistics>>();
const samples: PerformanceSample[] = [];
let completedRuns = 0;

function addSample(running: RunningStatistics, value: number): void {
	running.sampleCount += 1;
	const delta = value - running.mean;
	running.mean += delta / running.sampleCount;
	running.sumOfSquares += delta * (value - running.mean);
	running.min = Math.min(running.min, value);
	running.max = Math.max(running.max, value);
}

function recordSample(sample: PerformanceSample): void {
	samples.push(sample);
	let counters = statistics.get(sample.measureName);
	if (counters === undefined) {
		counters = new Map();
		statistics.set(sample.measureName, counters);
	}
	for (const counter of performanceCounters) {
		let running = counters.get(counter);
		if (running === undefined) {
			running = { sampleCount: 0, mean: 0, sumOfSquares: 0, min: Infinity, max: -Infinity };
			counters.set(counter, running);
		}
		addSample(running, sample[counter]);
	}
}

function getPerformanceStatistics(): PerformanceCounterStatistics[] {
	const result: PerformanceCounterStatistics[] = [];
	for (const [measureName, counters] of statistics) {
		for (const [counter, running] of counters) {
			result.push({
				backEnd,
				measureName,
				counter,
				sampleCount: running.sampleCount,
				mean: running.mean,
				standardDeviation:
					running.sampleCount > 1
						? Math.sqrt(running.sumOfSquares / (running.sampleCount - 1))
						: 0,
				min: running.min,
				max: running.max,
			});
		}
	}
	return result;
}

interface PerformanceMemory {
	readonly usedJSHeapSize: number;
}

/** Same check and order of operations as ScriptorPad's `#getWorkingSetAfterGarbageCollection`. */
function getWorkingSetAfterGarbageCollection(): number {
	const { gc } = window as unknown as { gc?: () => void };
	if (typeof gc !== "function") {
		throw new TypeError(
			'Working-set measurement requires Chromium launched with --js-flags="--expose-gc" and --enable-precise-memory-info.',
		);
	}
	gc();
	const memory = (performance as unknown as { memory?: PerformanceMemory }).memory;
	if (memory === undefined || !Number.isFinite(memory.usedJSHeapSize)) {
		throw new TypeError(
			"Working-set measurement requires performance.memory.usedJSHeapSize. Launch Chromium with --enable-precise-memory-info.",
		);
	}
	return memory.usedJSHeapSize;
}

// #endregion

// #region Editor lifecycle

/** The parts of the container's `IContainer` that the harness waits on. */
interface DirtyTrackingContainer {
	readonly isDirty: boolean;
	once(event: "saved", listener: () => void): void;
}

interface OpenEditor {
	readonly container: FluidContainer<TextEditorData>;
	readonly fluidContainer: DirtyTrackingContainer;
	readonly treeView: TreeViewAlpha<typeof TextEditorRoot>;
	readonly undoRedo: UndoRedo;
	readonly host: HTMLElement;
	readonly reactRoot: Root;
	readonly quill: Quill;
	readonly text: FormattedTextDefault.Tree;
}

if (new URLSearchParams(location.search).get("fluidClient") !== "ephemeral") {
	throw new Error("Load the performance harness with ?fluidClient=ephemeral.");
}
const client = getExampleServiceClient({ ...defaultServiceOptions });
const service = (client as unknown as { service: { deleteAllDocuments(): Promise<void> } })
	.service;

let currentEditor: OpenEditor | undefined;

/**
 * Closes the editor of the previous run and deletes all documents, which also replaces the in-memory server.
 */
async function closeCurrentEditor(): Promise<void> {
	const editor = currentEditor;
	if (editor === undefined) {
		return;
	}
	currentEditor = undefined;
	editor.reactRoot.unmount();
	editor.host.remove();
	editor.undoRedo.dispose();
	editor.treeView.dispose();
	editor.container.close();
	await service.deleteAllDocuments();
}

async function waitForQuill(host: HTMLElement): Promise<Quill> {
	const deadline = performance.now() + 30_000;
	for (;;) {
		const element = host.querySelector(".ql-container");
		const found = element === null ? null : Quill.find(element);
		if (found instanceof Quill) {
			return found;
		}
		if (performance.now() > deadline) {
			throw new Error("Timed out waiting for the Quill editor to mount.");
		}
		await new Promise((resolve) => setTimeout(resolve, 10));
	}
}

/** Creates a new attached document and mounts one formatted Quill editor on it, as the example app does. */
async function openEditor(): Promise<OpenEditor> {
	const container = await client.createAttachedContainer(TextEditorDataStore);
	const fluidContainer = (container as unknown as { container: DirtyTrackingContainer })
		.container;
	const treeView = container.data.treeView;
	const undoRedo = createUndoRedo(treeView);
	const host = document.createElement("div");
	host.style.width = "800px";
	host.style.height = "600px";
	document.body.append(host);
	const reactRoot = createRoot(host);
	const text = treeView.root.formattedText;
	reactRoot.render(<FormattedMainView root={toPropTreeNode(text)} undoRedo={undoRedo} />);
	const quill = await waitForQuill(host);
	return { container, fluidContainer, treeView, undoRedo, host, reactRoot, quill, text };
}

/**
 * Waits until the container has no unacknowledged local ops, then yields a macrotask.
 * Equivalent of ScriptorPad's `syncModelForSharedSessions`, which runs after every CommandScript command.
 */
async function settle(editor: OpenEditor): Promise<void> {
	let wasDirty: boolean;
	do {
		wasDirty = false;
		if (editor.fluidContainer.isDirty) {
			await new Promise<void>((resolve) => editor.fluidContainer.once("saved", () => resolve()));
			wasDirty = true;
		}
	} while (wasDirty);
	await new Promise((resolve) => setTimeout(resolve, 0));
}

/**
 * Called after every `PerformanceEnd`, outside the measured window, with the result just recorded. Lets a check
 * inspect the editor's DOM as it was at the end time.
 */
let measurementEndHook: ((measureName: string, sample: PerformanceSample) => void) | undefined;

/**
 * Runs `command` between the equivalents of `PerformanceStart` and `PerformanceEnd`, each followed by
 * {@link settle} like every CommandScript command. Without a command, it is a `PerformanceStart` directly followed by
 * a `PerformanceEnd`.
 */
async function measure(
	editor: OpenEditor,
	measureName: string,
	command?: () => void,
): Promise<void> {
	// PerformanceStart
	const initialWorkingSetBytes = getWorkingSetAfterGarbageCollection();
	const startTime = performance.now();
	await settle(editor);

	if (command !== undefined) {
		command();
		await settle(editor);
	}

	// PerformanceEnd
	const executionTimeMs = performance.now() - startTime;
	const finalWorkingSetBytes = getWorkingSetAfterGarbageCollection();
	const sample: PerformanceSample = {
		run: completedRuns,
		measureName,
		executionTimeMs,
		workingSetBytes: finalWorkingSetBytes - initialWorkingSetBytes,
		initialWorkingSetBytes,
		finalWorkingSetBytes,
	};
	recordSample(sample);
	measurementEndHook?.(measureName, sample);
	await settle(editor);
}

// #endregion

// #region Perf-Basic

/** The sample paragraphs of Scriptor's `=rand(n)` text generator (RandService). */
const sampleParagraphs: readonly string[] = [
	"Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.",
	"To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries.",
	"Themes and styles also help keep your document coordinated. When you click Design and choose a new Theme, the pictures, charts, and SmartArt graphics change to match your new theme. When you apply styles, your headings change to match the new theme.",
	"Save time in Word with new buttons that show up where you need them. To change the way a picture fits in your document, click it and a button for layout options appears next to it. When you work on a table, click where you want to add a row or a column, and then click the plus sign.",
	"Reading is easier, too, in the new Reading view. You can collapse parts of the document and focus on the text you want. If you need to stop reading before you reach the end, Word remembers where you left off - even on another device.",
];

/** The same text `=rand(n)` produces in Scriptor: numbered paragraphs, each ending with a paragraph break. */
function generateParagraphs(paragraphCount: number): string {
	let text = "";
	for (let index = 0; index < paragraphCount; index++) {
		text += `[${index}] ${sampleParagraphs[index % sampleParagraphs.length]}\n`;
	}
	return text;
}

function verify(condition: boolean, message: string): void {
	if (!condition) {
		throw new Error(`Verification failed: ${message}`);
	}
}

/** Selects all text, like Ctrl+A in the editor: everything but Quill's mandatory final line break. */
function selectAll(quill: Quill): void {
	quill.setSelection(0, quill.getLength() - 1, "user");
}

function getSelection(quill: Quill): { index: number; length: number } {
	const range = quill.getSelection(true);
	if (range === null) {
		throw new Error("The editor has no selection.");
	}
	return range;
}

/**
 * One run of Perf-Basic: open an editor, type `=rand(n)`, then measure generating the paragraphs, making all text
 * bold and deleting all text.
 */
/**
 * The first two commands of Perf-Basic, `openEditor` and typing `=rand(n)`, each followed by {@link settle}.
 * @returns The new editor and the insertion point after the typed text.
 */
async function openEditorAndTypeTrigger(
	paragraphCount: number,
): Promise<{ editor: OpenEditor; cursor: number }> {
	await closeCurrentEditor();

	// openEditor
	const editor = await openEditor();
	currentEditor = editor;
	const { quill, text } = editor;
	await settle(editor);

	// typing "=rand(n)": one command, and one transaction, per character.
	const trigger = `=rand(${paragraphCount})`;
	let cursor = 0;
	for (const character of trigger) {
		quill.insertText(cursor, character, "user");
		cursor += character.length;
		quill.setSelection(cursor, 0, "silent");
		await settle(editor);
	}
	verify(text.fullString() === trigger, `the document should contain "${trigger}"`);
	return { editor, cursor };
}

/**
 * The GenerateParagraphs step: Scriptor replaces the "=rand(n)" text with the generated paragraphs when Enter is
 * pressed, in one transaction, and leaves the insertion point in an empty paragraph after them.
 */
function generateParagraphsCommand(
	editor: OpenEditor,
	cursor: number,
	generated: string,
): () => void {
	return () => {
		editor.quill.updateContents([{ delete: cursor }, { insert: generated }], "user");
		editor.quill.setSelection(generated.length, 0, "silent");
	};
}

async function runPerfBasic(paragraphCount: number): Promise<void> {
	const { editor, cursor } = await openEditorAndTypeTrigger(paragraphCount);
	const { quill, text } = editor;

	const generated = generateParagraphs(paragraphCount);
	await measure(
		editor,
		"GenerateParagraphs",
		generateParagraphsCommand(editor, cursor, generated),
	);
	verify(text.fullString() === generated, "the document should contain the generated paragraphs");
	verify(quill.getText() === `${generated}\n`, "Quill should show the generated paragraphs");

	// selectAll
	selectAll(quill);
	await settle(editor);

	// BoldAll: what Quill's Bold toolbar button and Ctrl+B handlers do.
	await measure(editor, "BoldAll", () => {
		const range = getSelection(quill);
		quill.format("bold", quill.getFormat(range).bold !== true, "user");
	});
	{
		const content = text.fullString();
		const atoms = text.charactersWithFormatting();
		verify(atoms.length === content.length, "the document should only contain ASCII text");
		let textAtoms = 0;
		let boldAtoms = 0;
		for (const [index, atom] of atoms.entries()) {
			if (content[index] !== "\n") {
				textAtoms += 1;
				if (atom.format.bold) {
					boldAtoms += 1;
				}
			}
		}
		verify(textAtoms > 0 && boldAtoms === textAtoms, "all text should be bold");
	}

	// selectAll
	selectAll(quill);
	await settle(editor);

	// DeleteAll: what Quill's Delete/Backspace handlers do with a selection (keyboard module `deleteRange`).
	await measure(editor, "DeleteAll", () => {
		const range = getSelection(quill);
		quill.deleteText(range.index, range.length, "user");
		quill.setSelection(range.index, 0, "silent");
	});
	verify(text.characterCount() === 0, "the document should be empty");
	verify(quill.getText() === "\n", "Quill should show an empty document");

	completedRuns += 1;
}

/**
 * Calibration: a `PerformanceStart` directly followed by a `PerformanceEnd`, after the same `openEditor` and typing as
 * Perf-Basic. It shows the fixed cost of the measurement window itself, to compare with the same script in ScriptorPad.
 */
async function runEmptyMeasurement(): Promise<void> {
	const { editor } = await openEditorAndTypeTrigger(10);
	await measure(editor, "Empty");
	completedRuns += 1;
}

/** What the editor shows after the GenerateParagraphs step, for comparison with ScriptorPad's `=rand(n)`. */
interface GeneratedDocument {
	/** The text of every line (paragraph) of the Quill editor, without Quill's line breaks. */
	readonly quillParagraphs: string[];
	/** The text of the formatted text tree. */
	readonly treeText: string;
	/** The line that holds the insertion point, and the position of the insertion point in it. */
	readonly selectionParagraph: number;
	readonly selectionOffset: number;
	readonly selectionLength: number;
}

/** Runs Perf-Basic up to the GenerateParagraphs step, without measuring, and returns the resulting document. */
async function inspectGeneratedDocument(paragraphCount: number): Promise<GeneratedDocument> {
	const { editor, cursor } = await openEditorAndTypeTrigger(paragraphCount);
	generateParagraphsCommand(editor, cursor, generateParagraphs(paragraphCount))();
	await settle(editor);
	const { quill, text } = editor;
	const selection = getSelection(quill);
	const [line, offset] = quill.getLine(selection.index);
	return {
		// getText always ends with Quill's mandatory final line break, which ends the last line.
		quillParagraphs: quill.getText().slice(0, -1).split("\n"),
		treeText: text.fullString(),
		selectionParagraph: line === null ? -1 : quill.getLines().indexOf(line),
		selectionOffset: offset,
		selectionLength: selection.length,
	};
}

// #endregion

/** The automation API used by the Playwright sweep (perf/quillPerfScaling.mjs). */
interface QuillPerfApi {
	/**
	 * Runs a scenario `repeatCount` times: Perf-Basic (the default) or the empty calibration measurement.
	 */
	run(request: {
		paragraphCount: number;
		repeatCount: number;
		scenario?: "Perf-Basic" | "Empty";
	}): Promise<{ status: "passed" | "failed"; error?: string }>;
	resetPerformanceStatistics(): void;
	getPerformanceStatistics(): PerformanceCounterStatistics[];
	getSamples(): PerformanceSample[];
	/** The paragraphs `=rand(n)` generates in Scriptor, the last being the empty one with the insertion point. */
	expectedParagraphs(paragraphCount: number): string[];
	inspect(paragraphCount: number): Promise<GeneratedDocument>;
	/** Sets the function called after every `PerformanceEnd`, outside the measured window. */
	setMeasurementEndHook(
		hook: ((measureName: string, sample: PerformanceSample) => void) | undefined,
	): void;
}

let queue: Promise<unknown> = Promise.resolve();

const api: QuillPerfApi = {
	async run({ paragraphCount, repeatCount, scenario = "Perf-Basic" }) {
		const execution = queue.then(async () => {
			try {
				for (let repeat = 0; repeat < repeatCount; repeat++) {
					await (scenario === "Empty"
						? runEmptyMeasurement()
						: runPerfBasic(paragraphCount));
				}
				return { status: "passed" as const };
			} catch (error) {
				console.error(error);
				return {
					status: "failed" as const,
					error: error instanceof Error ? `${error.message}\n${error.stack}` : String(error),
				};
			}
		});
		queue = execution;
		return execution;
	},
	resetPerformanceStatistics() {
		statistics.clear();
		samples.length = 0;
	},
	getPerformanceStatistics,
	getSamples: () => [...samples],
	expectedParagraphs: (paragraphCount) => generateParagraphs(paragraphCount).split("\n"),
	async inspect(paragraphCount) {
		const execution = queue.then(async () => inspectGeneratedDocument(paragraphCount));
		queue = execution.catch(() => {});
		return execution;
	},
	setMeasurementEndHook(hook) {
		measurementEndHook = hook;
	},
};

(window as unknown as { quillPerf: QuillPerfApi }).quillPerf = api;
