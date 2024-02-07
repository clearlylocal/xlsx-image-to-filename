import { Command } from 'cliffy/command/mod.ts'
import { convert } from './convert.ts'
import { toDefaultOutputFilePath } from './utils.ts'

type Params = {
	inPath: string
	outPath: string
	column?: string
}

export async function cli() {
	await new Command()
		.name('conditionalize-pptx')
		.version('0.1.0')
		.description('Conditional content for PPTs')
		.option('-f, --file-path <file>', 'Input file path', { required: true })
		.option('-o, --out-path <file>', 'Output file path', { required: false })
		.option('-c, --column <string>', 'Column to use for output', { required: false })
		.action(async ({ filePath, outPath, column }, ..._args) => {
			outPath ??= toDefaultOutputFilePath(filePath)
			await run({
				inPath: filePath,
				outPath,
				column,
			})
		})
		.parse(Deno.args)
}

async function run({ inPath, outPath, column }: Params) {
	const bytes = await Deno.readFile(inPath)
	const outBytes = await convert(bytes, column)

	await Deno.writeFile(outPath, outBytes)
	console.info(`Wrote to ${outPath}`)
}
