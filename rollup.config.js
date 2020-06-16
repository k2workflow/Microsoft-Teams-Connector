import resolve from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import replace from '@rollup/plugin-replace';
import json from '@rollup/plugin-json';
import babel from 'rollup-plugin-babel';
import { terser } from "rollup-plugin-terser";

const BUILD_ENV = JSON.stringify(process.env.NODE_ENV || 'development');
const devBuild = process.env.NODE_ENV === 'development';

function getOutputFile(inputFile) {
    const inputFileIndex = inputFile.lastIndexOf('.');
    const outputFile = (inputFileIndex >= 0 ? inputFile.substr(0, inputFileIndex) : inputFile) + ".js";
    return (outputFile.startsWith("src/"))
        ? outputFile.substr(4)
        : outputFile;
}

function buildConfig(inputFile) {

    const outputFile = getOutputFile(inputFile);
    const extensions = ['.js', '.ts'];
    const runtimeHelpers = true;

    const plugins = [
        resolve(),
        commonjs(),
        replace({ BUILD_ENV, 'process.env.NODE_ENV': BUILD_ENV }),
        json(),
        babel({ extensions, runtimeHelpers })
    ];

    if (!devBuild) {
        plugins.push(
            terser()
        );
    }

    return {
        input: inputFile,
        output: {
            file: 'dist/' + outputFile,
            format: 'iife',
            sourcemap: true,
            strict: false, // HACK: Only required for preview.
        },
        plugins
    };
}

function buildTestConfig(inputFile) {

    const outputFile = getOutputFile(inputFile);
    const extensions = ['.js', '.ts'];
    const runtimeHelpers = true;

    const plugins = [
        resolve(),
        commonjs(),
        replace({ BUILD_ENV, 'process.env.NODE_ENV': BUILD_ENV }),
        json(),
        babel({ extensions, runtimeHelpers })
    ];

    if (!devBuild) {
        plugins.push(
            terser()
        );
    }

    return {
        input: inputFile,
        external: [ "ava" ],
        external: () => true,
        output: {
            file: 'dist/' + outputFile,
            format: 'cjs',
            sourcemap: true
        },
        plugins
    };
}

export default [
    buildConfig('src/index.ts'),
    buildTestConfig('src/test.ts')
];