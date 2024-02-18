import { colors } from 'cliffy/ansi/colors.ts'
import { Command } from 'cliffy/command/mod.ts'
import { convert } from './convert.ts'
import { expandGlob } from 'std/fs/expand_glob.ts'
// TODO
import { type WalkEntry } from 'std/fs/_util.ts'
import { common, dirname, join, relative } from 'std/path/mod.ts'

const IS_COMPILED_FLAG = '--__is-compiled'
export const IS_COMPILED = Deno.args.includes(IS_COMPILED_FLAG)

const cmdName = 'xlsx-image-to-filename'

const numFormat = new Intl.NumberFormat('en-US')

function fmtFilePath(filePath: string) {
	return `${colors.gray('<')}${relative('.', filePath)}${colors.gray('>')}`
}

export async function cli() {
	await new Command()
		.name(cmdName)
		.version('0.2.0')
		.description('Add file names to XLSX files based on images present in each row')
		.option(
			'-p, --prefix <string>',
			'Prefix to add before file name, e.g. "https://clearlyloc.sharepoint.com/sites/ProjectScreenshots/oss1/" (default: "")',
			{ required: false },
		)
		.option('-c, --column <string>', 'Column to use for output (default: "O")', { required: false })
		.option(
			'-o, --out-path <string>',
			'Output file path (default: input file path with "_with_image_file_names_{{DATE_TIME}}" appended)',
			{ required: false },
		)
		.arguments('<input_path:string> [...other_input_paths:string]')
		.action(
			async (params, ...paths) => {
				const entries: WalkEntry[] = []
				for (const path of paths) {
					for await (const entry of expandGlob(path)) {
						entries.push(entry)
					}
				}
				if (entries.length === 1 && entries[0].isDirectory) {
					for await (const entry of expandGlob(join(entries[0].path, '*'))) {
						entries.push(entry)
					}
				}

				const fileEntries = entries.filter((entry) => entry.isFile && entry.name.endsWith('.xlsx'))
				const commonPrefix = common(entries.map((entry) => entry.path))

				console.info(
					`\nFound ${colors.cyan.bold(numFormat.format(fileEntries.length))} matching XLSX files in ${
						fmtFilePath(commonPrefix)
					}`,
				)
				console.info(`Writing...`)

				const digitsLength = numFormat.format(fileEntries.length).length

				let numWritten = 0
				const fns = fileEntries.map((entry) => async () => {
					const bytes = await Deno.readFile(entry.path)
					const out = await convert(bytes, { ...params, filePath: entry.path })

					try {
						await Deno.stat(dirname(out.path))
					} catch {
						await Deno.mkdir(dirname(out.path), { recursive: true })
					}
					await Deno.writeFile(out.path, out.bytes)
					numWritten++

					const n = numFormat.format(numWritten).padStart(digitsLength, ' ')

					console.info(
						`${colors.gray(`[${n}/${numFormat.format(fileEntries.length)}]`)} Wrote ${
							fmtFilePath(relative(commonPrefix, entry.path))
						} ⇒ ${colors.cyan(fmtFilePath(out.path))}`,
					)
				})

				if (IS_COMPILED) {
					// run in serial
					for (const fn of fns) await fn()
				} else {
					// run in parallel
					await Promise.all(fns.map((fn) => fn()))
				}

				console.info(`Done!`)
			},
		)
		.example(
			'Basic example',
			`${cmdName} input_files --out-path ${
				join('output_files', '{{DATE}}')
			} --prefix https://clearlyloc.sharepoint.com/sites/SiteName/path/to/library`,
		)
		.parse(Deno.args.filter((a) => a !== IS_COMPILED_FLAG))
}
