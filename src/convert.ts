import { dirname, join as posixJoin } from 'std/path/posix/mod.ts'
import { BlobReader, BlobWriter, ZipReader, ZipWriter } from 'zipjs'
import {
	dateForFilePath,
	dateTimeForFilePath,
	expandRange,
	get$,
	getRightMostRecord,
	normalizePathForCurrentOs,
	posixifyPath,
	ptsToEmus,
	toCellReferences,
	toDefaultOutputFilePath,
	toRelsPath,
} from './utils.ts'
import type { Entry, Params } from './types.ts'
import { colors } from 'cliffy/ansi/colors.ts'
import type { Cheerio, Element } from 'cheerio'
import { letterToColIndex } from './utils.ts'
import { IS_COMPILED } from './cli.ts'
import { join, SEP } from 'std/path/mod.ts'

export async function* convertAll(files: Uint8Array[], params: Params) {
	for (const file of files) {
		yield convert(file, params)
	}
}

type ImageReference = {
	cellReference: string
	kind: 'floating' | 'wps'
	fileName: string
}

class EntryNotFoundError extends Error {}

export async function convert(fileBytes: Uint8Array, { column, prefix, filePath, outPath }: Params) {
	const outputCol = column ??= 'O'
	prefix ??= ''

	const fileName = filePath.split(SEP).at(-1)!.split('.').at(0)!

	prefix = prefix.replaceAll('{{FILE_NAME}}', encodeURIComponent(fileName))
	if (prefix && prefix.includes('/') && !prefix.endsWith('/')) prefix += '/'

	outPath ??= toDefaultOutputFilePath(filePath)

	outPath = normalizePathForCurrentOs(outPath)

	outPath = outPath
		.replaceAll('{{DATE}}', dateForFilePath())
		.replaceAll('{{DATE_TIME}}', dateTimeForFilePath())
		.replaceAll('{{FILE_NAME}}', fileName)

	if (!outPath.split(SEP).includes(fileName)) {
		outPath = join(outPath, `${fileName}.xlsx`)
	}

	const warnings: string[] = []

	const blob = new Blob([fileBytes])

	const blobReader = new BlobReader(blob)
	const zipReader = new ZipReader(blobReader, {
		useWebWorkers: !IS_COMPILED,
	})

	const blobWriter = new BlobWriter()
	const zipWriter = new ZipWriter(blobWriter)

	const entries = await zipReader.getEntries() as Entry[]

	function get$FromPath(path: string) {
		const entry = entries.find((x) => posixifyPath(x.filename) === posixifyPath(path))
		if (!entry) {
			throw new EntryNotFoundError(`Entry for path ${path} not found`)
		}

		return get$(entry)
	}

	const cellImagesPath = 'xl/cellimages.xml'
	const cellImageRelsPath = toRelsPath(cellImagesPath)

	const hasWpsCellImages = entries.some((x) => posixifyPath(x.filename) === posixifyPath(cellImagesPath))

	const pathsAlreadyWritten: string[] = []

	{
		const wpsIdsToImageFileNames = new Map<string, string>()

		if (hasWpsCellImages) {
			const _ridsToWpsIds = new Map<string, string[]>()

			{
				const $ = await get$FromPath(cellImagesPath)
				const cellImages = $('etc\\:cellImage')

				for (const img of cellImages) {
					const $img = $(img)
					const name = $img.find('xdr\\:cNvPr').attr('name')
					const rid = $img.find('xdr\\:blipFill a\\:blip').attr('r:embed')

					if (!rid || !name) continue

					const wpsIds = _ridsToWpsIds.get(rid) ?? []
					wpsIds.push(name)

					_ridsToWpsIds.set(rid, wpsIds)
				}
			}

			{
				const $ = await get$FromPath(cellImageRelsPath)
				const imageRels = $(
					'Relationship[Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"]',
				)

				for (const rel of imageRels) {
					const $rel = $(rel)

					const rid = $rel.attr('Id')

					if (!rid) continue

					const wpsIds = _ridsToWpsIds.get(rid)
					const target = $rel.attr('Target')?.split('/').at(-1)

					if (!wpsIds || !target) continue

					for (const wpsId of wpsIds) {
						wpsIdsToImageFileNames.set(wpsId, target)
					}
				}
			}
		}

		const sheetEntries = entries.filter((x) => /^xl\/worksheets\/\w+.xml$/.test(x.filename))

		for (const sheetEntry of sheetEntries) {
			const relsFilePath = toRelsPath(sheetEntry.filename)
			const imageReferences: ImageReference[] = []

			// WPS

			const $ = await get$(sheetEntry)

			const $cells = $('sheetData row c')

			for (const cell of $cells) {
				const $cell = $(cell)
				const $f = $cell.find('f, v')
				if (!$f.length) continue

				const formula = $f.text()

				const match = formula.match(/\bDISPIMG\("([^"]+)/)

				if (match) {
					const id = match[1]
					const imgFilename = wpsIdsToImageFileNames.get(id)

					if (imgFilename) {
						imageReferences.push({
							cellReference: $cell.attr('r')!,
							kind: 'wps',
							fileName: imgFilename,
						})
					} else {
						warnings.push(`Image file for ${id} not found`)
					}
				}
			}

			// repeat WPS for merged cells (multi-cell floating images already covered by positioning logic)

			for (const mergeCell of $('mergeCells mergeCell')) {
				// e.g. <mergeCell ref="A16:A17"/>
				const [start, end] = $(mergeCell).attr('ref')!.split(':')

				const imageReference = imageReferences.find((x) => x.cellReference === start)
				if (imageReference) {
					for (const cellReference of expandRange(start, end).slice(1)) {
						const newImgRef = {
							...imageReference,
							cellReference,
						}

						imageReferences.push(newImgRef)
					}
				}
			}

			// floating images

			const relsEntry = entries.find((x) => posixifyPath(x.filename) === posixifyPath(relsFilePath))

			if (relsEntry) {
				const colBoundaries: number[] = [0]
				const rowBoundaries: number[] = [0]

				// TODO: col width still not accurate vs emus
				// figure out later as shouldn't affect end calc
				let c = 0
				for (const col of $('cols col')) {
					const width = Number($(col).attr('width'))
					c += width
					// colBoundaries.push(excelWidthToEmus(c))
					colBoundaries.push(ptsToEmus(c))
				}

				let r = 0
				for (const row of $('sheetData row')) {
					const height = Number($(row).attr('ht'))
					r += height
					rowBoundaries.push(ptsToEmus(r))
				}

				const toNearestCellReferences = toCellReferences({ colBoundaries, rowBoundaries })

				{
					const $ = await get$(relsEntry)

					const drawingsPathRelative = $(
						'Relationship[Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"]',
					).attr('Target')

					if (drawingsPathRelative) {
						const drawingsPath = posixJoin(dirname(relsFilePath), '..', drawingsPathRelative)
						const drawingRelsPath = toRelsPath(drawingsPath)

						const $ = await get$FromPath(drawingsPath)
						const rels$ = await get$FromPath(drawingRelsPath)

						for (const anchor of $('xdr\\:twoCellAnchor')) {
							const $anchor = $(anchor)

							const rid = $anchor.find('xdr\\:blipFill a\\:blip').attr('r:embed')

							const $xfrm = $anchor.find('xdr\\:spPr a\\:xfrm')
							const $off = $xfrm.find('a\\:off')
							const $ext = $xfrm.find('a\\:ext')

							const fileName = rels$(`Relationship[Id="${rid}"]`).attr('Target')!.split('/').at(-1)!

							const coords = {
								x: Number($off.attr('x')),
								y: Number($off.attr('y')),
								cx: Number($ext.attr('cx')),
								cy: Number($ext.attr('cy')),
							}

							for (
								const cellReference of toNearestCellReferences(coords)
							) {
								const imageReference = {
									cellReference,
									kind: 'floating' as const,
									fileName,
								}

								imageReferences.push(imageReference)
							}
						}
					}
				}
			}

			// write rows into XML in column O
			const imageReferencesByRow = Object.groupBy(imageReferences, (x) => x.cellReference.match(/\d+/)![0])

			for (const [_row, _records] of Object.entries(imageReferencesByRow)) {
				const rowNum = Number(_row)
				const record = getRightMostRecord(_records!)
				if (!record) continue
				const $row = $(`sheetData row[r="${rowNum}"]`)
				const outputCellRef = `${outputCol}${rowNum}`

				const cellSelector = `c[r="${outputCellRef}"]`
				let $cell = $row.find(cellSelector)

				append: if (!$cell.length) {
					const $toAppend = $(`<c r="${outputCellRef}"></c>`) as Cheerio<Element>

					for (const c of $row.find('c')) {
						const $c = $(c)

						if (letterToColIndex($c.attr('r')!.match(/\D+/)![0]) > letterToColIndex(outputCol)) {
							$toAppend.insertBefore($c)
							break append
						}
					}

					$row.append($toAppend)
				}

				$cell = $row.find(cellSelector)

				// for (const attr of $cell.prop('attributes')) {
				// 	$cell.removeAttr(attr.name)
				// }

				$cell.attr('t', 'str')

				$cell.html('<v></v>')
				$cell.find('v').text(`${prefix}${record.fileName}`)
			}

			// write to ZIP
			pathsAlreadyWritten.push(sheetEntry.filename)
			zipWriter.add(sheetEntry.filename, new Blob([$.xml()]).stream(), { useWebWorkers: !IS_COMPILED })
		}
	}

	for (const entry of entries) {
		if (pathsAlreadyWritten.includes(entry.filename as typeof pathsAlreadyWritten[number])) {
			continue
		}

		zipWriter.add(
			entry.filename,
			(await entry.getData(new BlobWriter(), { useWebWorkers: !IS_COMPILED })).stream(),
			{
				useWebWorkers: !IS_COMPILED,
			},
		)
	}

	if (warnings.length) {
		for (const warning of warnings) {
			console.warn(colors.bgYellow(warning))
		}
	}

	return {
		path: outPath,
		bytes: new Uint8Array(await (await zipWriter.close()).arrayBuffer()),
	}
}
