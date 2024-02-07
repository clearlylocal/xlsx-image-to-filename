import { Command } from 'cliffy/command/mod.ts'
import { convert } from './convert.ts'
import { toDefaultOutputFilePath } from './utils.ts'

const IS_COMPILED_FLAG = '--__is-compiled'
export const IS_COMPILED = Deno.args.includes(IS_COMPILED_FLAG)

export async function cli() {
	await new Command()
		.name('conditionalize-pptx')
		.description('Conditional content for PPTs')
		.option('-f, --file-path <file>', 'Input file path', { required: true })
		.option(
			'-p, --prefix <string>',
			'Prefix to add before file name, e.g. "https://clearlyloc.sharepoint.com/sites/ProjectScreenshots/oss1/" (default: "")',
			{ required: false },
		)
		.option('-c, --column <string>', 'Column to use for output (default: "O")', { required: false })
		.option(
			'-o, --out-path <file>',
			'Output file path (default: input file path with "_with_image_file_names_{{DATE}}" appended)',
			{ required: false },
		)
		.action(
			async (params) => {
				let { filePath, outPath } = params
				const bytes = await Deno.readFile(filePath)
				const outBytes = await convert(bytes, params)

				outPath ??= toDefaultOutputFilePath(filePath)

				await Deno.writeFile(outPath, outBytes)
				console.info(`Wrote to ${outPath}`)
			},
		)
		.example(
			'Basic example',
			'xlsx-image-to-filename --file-path "oss图文对照表1.xlsx" --prefix "https://clearlyloc.sharepoint.com/sites/ProjectScreenshots/oss1/"',
		)
		.parse(Deno.args.filter((a) => a !== IS_COMPILED_FLAG))
}
