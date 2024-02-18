// JS port from https://stackoverflow.com/questions/48984697/convert-a-number-to-excel-s-base-26/
function divmodExcel(n: number) {
	const a = Math.floor(n / 26)
	const b = n % 26

	return b === 0 ? [a - 1, b + 26] : [a, b]
}

const uppercaseAlpha = Array.from({ length: 26 }, (_, i) => String.fromCodePoint(i + 'A'.codePointAt(0)!))

/** @param n The 1-indexed Excel column number */
export function toExcelCol(n: number) {
	const chars = []

	let d: number
	while (n > 0) {
		;[n, d] = divmodExcel(n)
		chars.unshift(uppercaseAlpha[d - 1])
	}
	return chars.join('')
}

/** @returns The 1-indexed Excel column number */
export function fromExcelCol(letter: string) {
	return [...letter].map((l) => uppercaseAlpha.indexOf(l)).reduce((acc, cur) => {
		return acc * 26 + cur + 1
	}, 0)
}
