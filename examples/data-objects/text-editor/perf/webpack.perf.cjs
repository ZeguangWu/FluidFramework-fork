/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

// Builds the Quill-on-SharedTree performance harness (perf/quillPerfHarness.tsx) into perf/dist.
// It uses development mode by default, like the ScriptorPad bundle it is compared with.
// Pass `--env production` for an optimized build.

const path = require("node:path");

const fluidRoute = require("@fluid-example/webpack-fluid-loader");

module.exports = (env = {}) => {
	const packageDirectory = path.resolve(__dirname, "..");
	const config = fluidRoute.exampleAppConfig(packageDirectory, env, {
		html: { title: "Quill on SharedTree performance harness" },
	});
	const typescriptRule = config.module.rules.find(
		(rule) => String(rule.test) === String(/\.tsx?$/),
	);
	if (typescriptRule === undefined) {
		throw new Error("The example configuration has no TypeScript rule.");
	}

	return {
		...config,
		context: packageDirectory,
		entry: {
			main: "./perf/quillPerfHarness.tsx",
		},
		output: {
			...config.output,
			path: path.resolve(__dirname, "dist"),
		},
		// The harness bundle has no source maps, like the ScriptorPad bundle.
		devtool: false,
		module: {
			rules: [
				...config.module.rules.filter((rule) => rule !== typescriptRule),
				{
					test: /\.tsx?$/,
					loader: typescriptRule.loader,
					options: {
						transpileOnly: true,
						configFile: path.resolve(__dirname, "tsconfig.json"),
					},
				},
				{
					test: /\.css$/,
					use: ["style-loader", "css-loader"],
				},
			],
		},
		performance: {
			hints: false,
		},
	};
};
