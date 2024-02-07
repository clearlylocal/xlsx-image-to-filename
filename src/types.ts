import type { Entry as _Entry } from 'zipjs'

// https://github.com/gildas-lormeau/zip.js/issues/371
export type Entry = _Entry & { getData: Exclude<_Entry['getData'], undefined> }

export type Params = {
	filePath: string
	outPath?: string
	column?: string
	prefix?: string
	isCompiled?: boolean
}
