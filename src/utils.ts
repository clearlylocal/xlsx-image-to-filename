import { load } from 'cheerio'
import type { Entry } from './types.ts'
import { TextWriter } from 'zipjs'
import { SEP as POSIX_SEP } from 'std/path/posix/mod.ts'
import { SEP_PATTERN as WIN_SEP_PATTERN } from 'std/path/windows/mod.ts'
import { IS_COMPILED } from './cli.ts'
import { fromExcelCol, toExcelCol } from './columnLetters.ts'

export function posixifyPath(path: string) {
	return path.replaceAll(new RegExp(WIN_SEP_PATTERN.source, 'g'), POSIX_SEP)
}

export function excelWidthToEmus(width: number) {
	// https://stackoverflow.com/a/17930457/8318731
	// [Width in PT] = ([width in excel] / 10 * 70 + 5) / 96 * 72)
	// For example: Width 10 in excel = 75 pixels = 56.25 pts = .78125"

	const pts = (width / 10 * 70 + 5) / 96 * 72

	return ptsToEmus(pts)
}

export function ptsToEmus(width: number) {
	// 1pt = 12700 EMUs
	return width * 12_700
}

function seekClosestIndex(n: number, xs: number[]) {
	if (!xs.length) return -1
	if (n < xs[0]) return 0
	if (n > xs.at(-1)!) return xs.length - 1

	let lastErrorMargin = 0

	for (const [i, x] of xs.entries()) {
		const currentErrorMargin = n - x
		if (
			currentErrorMargin <= 0
		) {
			return Math.abs(lastErrorMargin) >= Math.abs(currentErrorMargin) ? i : i - 1
		}

		lastErrorMargin = currentErrorMargin
	}

	return xs.length - 1
}

function cellCoordsToRef([x, y]: readonly [x: number, y: number]) {
	return `${colIndexToLetter(x)}${y + 1}`
}

function formatToReferences(
	{ startColIdx, startRowIdx, endColIdx, endRowIdx }: {
		startColIdx: number
		startRowIdx: number
		endColIdx: number
		endRowIdx: number
	},
) {
	return (Array.from({ length: endColIdx - startColIdx }, (_, x) => x + startColIdx)
		.flatMap((x) => Array.from({ length: endRowIdx - startRowIdx }, (_, y) => [x, y + startRowIdx] as const)))
		.map(cellCoordsToRef)
}

export function toCellReferences(
	{ colBoundaries, rowBoundaries }: { colBoundaries: number[]; rowBoundaries: number[] },
) {
	return ({ x, y, cx, cy }: { x: number; y: number; cx: number; cy: number }) => {
		const startColIdx = seekClosestIndex(x, colBoundaries)
		const startRowIdx = seekClosestIndex(y, rowBoundaries)
		const endColIdx = seekClosestIndex(x + cx, colBoundaries) + 1
		const endRowIdx = seekClosestIndex(y + cy, rowBoundaries)

		return formatToReferences({ startColIdx, startRowIdx, endColIdx, endRowIdx })
	}
}

export function getRightMostRecord<T extends unknown>(records: T[]) {
	// TODO - care about col # or duplicates? probably not
	return records.at(-1)
}

export function expandRange(start: string, end: string) {
	const startColIdx = letterToColIndex(start.match(/\D+/)![0])
	const startRowIdx = Number(start.match(/\d+/)![0]) - 1
	const endColIdx = letterToColIndex(end.match(/\D+/)![0]) + 1
	const endRowIdx = Number(end.match(/\d+/)![0])

	return formatToReferences({ startColIdx, startRowIdx, endColIdx, endRowIdx })
}

export function colIndexToLetter(colIndex: number) {
	return toExcelCol(colIndex + 1)
}

export function letterToColIndex(letter: string) {
	return fromExcelCol(letter) - 1
}

export function toRelsPath(path: string) {
	return path.replace(/([^/\\]+)\.xml$/, '_rels/$1.xml.rels')
}

export async function get$(entry: Entry) {
	const tw = new TextWriter()
	const cellImagesContent = await entry.getData(tw, { useWebWorkers: !IS_COMPILED })

	return load(cellImagesContent, { xml: true })
}

export function toDefaultOutputFilePath(inputFilePath: string) {
	const replacement = `_with_image_file_names_${
		new Date().toISOString().replace(/\..+$/, '').replaceAll(/\D+/g, '')
	}.xlsx` as const

	return inputFilePath.replace(
		/(\.xlsx)?$/,
		replacement,
	) as `${string}${typeof replacement}`
}
