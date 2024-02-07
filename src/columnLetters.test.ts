// See https://stackoverflow.com/questions/48984697/convert-a-number-to-excel-s-base-26/

import { assertEquals } from 'std/assert/mod.ts'
import { fromExcelCol, toExcelCol } from './columnLetters.ts'

Deno.test('excel col letters', async (t) => {
	await t.step('toExcelCol', async (t) => {
		await t.step('toExcelCol(26) == "Z"', () => assertEquals(toExcelCol(26), 'Z'))
		await t.step('toExcelCol(27) == "AA"', () => assertEquals(toExcelCol(27), 'AA'))
		await t.step('toExcelCol(702) == "ZZ"', () => assertEquals(toExcelCol(702), 'ZZ'))
		await t.step('toExcelCol(703) == "AAA"', () => assertEquals(toExcelCol(703), 'AAA'))
	})

	await t.step('fromExcelCol', async (t) => {
		await t.step('fromExcelCol("Z") == 26', () => assertEquals(fromExcelCol('Z'), 26))
		await t.step('fromExcelCol("AA") == 27', () => assertEquals(fromExcelCol('AA'), 27))
		await t.step('fromExcelCol("ZZ") == 702', () => assertEquals(fromExcelCol('ZZ'), 702))
		await t.step('fromExcelCol("AAA") == 703', () => assertEquals(fromExcelCol('AAA'), 703))
	})
})
