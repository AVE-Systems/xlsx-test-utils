<?php

namespace Tests;

use AveSystems\XlsxTestUtils\XlsxAssertsTrait;
use const DIRECTORY_SEPARATOR;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PHPUnit\Framework\ExpectationFailedException;
use PHPUnit\Framework\TestCase;

/**
 * @internal
 * @coversNothing
 */
class XlsxAssertsTraitTest extends TestCase
{
    use XlsxAssertsTrait;

    public function testAssertXlsxCellValueEqualsSuccessful()
    {
        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellValueEquals('mixed font', $sheet, 'A1');
        $this->assertXlsxCellValueEquals('regular font', $sheet, 'A2');
        $this->assertXlsxCellValueEquals('green', $sheet, 'B1');
        $this->assertXlsxCellValueEquals('red', $sheet, 'B2');
        $this->assertXlsxCellValueEquals('yellow', $sheet, 'B3');
    }

    public function testAssertXlsxCellValueEqualsShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.A1 cell value does not equal expected value.'
        );

        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellValueEquals('mixed FONT', $sheet, 'A1');
    }

    public function testAssertXlsxCellEmpty()
    {
        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellEmpty($sheet, 'A3');
    }

    public function testAssertXlsxCellEmptyShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.B1 cell value is not empty.'
        );

        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellEmpty($sheet, 'B1');
    }

    public function testAssertXlsxCellFontColorEquals()
    {
        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellFontColorEquals(
            Color::COLOR_BLACK,
            $sheet,
            'A1'
        );
        $this->assertXlsxCellFontColorEquals(
            Color::COLOR_BLACK,
            $sheet,
            'A2'
        );
        $this->assertXlsxCellFontColorEquals(
            Color::COLOR_GREEN,
            $sheet,
            'B1'
        );
        $this->assertXlsxCellFontColorEquals(
            Color::COLOR_RED,
            $sheet,
            'B2'
        );
        $this->assertXlsxCellFontColorEquals(
            Color::COLOR_YELLOW,
            $sheet,
            'B3'
        );
    }

    public function testAssertXlsxCellFontColorEqualsShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.A1 cell font color does not equal expected value.'
        );

        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellFontColorEquals(
            Color::COLOR_GREEN,
            $sheet,
            'A1'
        );
    }

    public function testAssertXlsxCellBackgroundColorEquals()
    {
        // |A    |B     |C     |D     |E     |
        //1|green|green |green |      |      |
        //2|     |yellow|yellow|yellow|yellow|
        //3|red  |red   |red   |red   |blue  |
        $sheet = $this->loadSheet('example_background.xlsx');

        $this->assertXlsxCellBackgroundColorEquals(
            Color::COLOR_GREEN,
            Color::COLOR_GREEN,
            $sheet,
            'A1'
        );

        $this->assertXlsxCellBackgroundColorEquals(
            Color::COLOR_YELLOW,
            Color::COLOR_YELLOW,
            $sheet,
            'B2'
        );

        $this->assertXlsxCellBackgroundColorEquals(
            Color::COLOR_RED,
            Color::COLOR_RED,
            $sheet,
            'C3'
        );

        $this->assertXlsxCellBackgroundColorEquals(
            Color::COLOR_BLUE,
            Color::COLOR_BLUE,
            $sheet,
            'E3'
        );
    }

    public function testAssertXlsxCellBackgroundColorEqualsShouldThrowException()
    {
        // |A    |B     |C     |D     |E     |
        //1|green|green |green |      |      |
        //2|     |yellow|yellow|yellow|yellow|
        //3|red  |red   |red   |red   |blue  |
        $sheet = $this->loadSheet('example_background.xlsx');

        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.D3 cell background start color does not equal expected value.'
        );

        $this->assertXlsxCellBackgroundColorEquals(
            Color::COLOR_YELLOW,
            Color::COLOR_YELLOW,
            $sheet,
            'D3'
        );
    }

    public function testAssertXlsxCellsBackgroundColorEquals()
    {
        // |A    |B     |C     |D     |E     |
        //1|green|green |green |      |      |
        //2|     |yellow|yellow|yellow|yellow|
        //3|red  |red   |red   |red   |blue  |
        $sheet = $this->loadSheet('example_background.xlsx');

        $this->assertXlsxCellsBackgroundColorEquals(
            Color::COLOR_GREEN,
            Color::COLOR_GREEN,
            $sheet,
            'A1:C1'
        );

        $this->assertXlsxCellsBackgroundColorEquals(
            Color::COLOR_YELLOW,
            Color::COLOR_YELLOW,
            $sheet,
            'B2:E2'
        );

        $this->assertXlsxCellsBackgroundColorEquals(
            Color::COLOR_RED,
            Color::COLOR_RED,
            $sheet,
            'A3:D3'
        );

        $this->assertXlsxCellsBackgroundColorEquals(
            Color::COLOR_BLUE,
            Color::COLOR_BLUE,
            $sheet,
            'E3:E3'
        );
    }

    public function testAssertXlsxCellsBackgroundColorEqualsShouldThrowException()
    {
        // |A    |B     |C     |D     |E     |
        //1|green|green |green |      |      |
        //2|     |yellow|yellow|yellow|yellow|
        //3|red  |red   |red   |red   |blue  |
        $sheet = $this->loadSheet('example_background.xlsx');

        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.E3 cell background start color does not equal expected value.'
        );

        $this->assertXlsxCellsBackgroundColorEquals(
            Color::COLOR_RED,
            Color::COLOR_RED,
            $sheet,
            'A3:E3'
        );
    }

    public function testAssertXlsxCellHorizontalAlignmentEquals()
    {
        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellHorizontalAlignmentEquals(
            Alignment::HORIZONTAL_CENTER,
            $sheet,
            'C1'
        );
    }

    public function testAssertXlsxCellHorizontalAlignmentEqualsShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.C2 cell horizontal alignment does not equal expected value.'
        );

        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellHorizontalAlignmentEquals(
            Alignment::HORIZONTAL_CENTER,
            $sheet,
            'C2'
        );
    }

    public function testAssertXlsxCellVerticalAlignmentEquals()
    {
        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellVerticalAlignmentEquals(
            Alignment::VERTICAL_CENTER,
            $sheet,
            'C2'
        );
    }

    public function testAssertXlsxCellVerticalAlignmentEqualsShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.C1 cell vertical alignment does not equal expected value.'
        );

        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellVerticalAlignmentEquals(
            Alignment::VERTICAL_CENTER,
            $sheet,
            'C1'
        );
    }

    public function testAssertXlsxCellWrapTextAlignmentTrue()
    {
        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellWrapTextAlignmentTrue(
            $sheet,
            'D1'
        );
    }

    public function testAssertXlsxCellWrapTextAlignmentTrueShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.C1 cell does not have wrap text.'
        );

        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellWrapTextAlignmentTrue(
            $sheet,
            'C1'
        );
    }

    public function testAssertXlsxColumnWidthEquals()
    {
        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxColumnWidthEquals(
            57.29,
            $sheet,
            'C'
        );
        $this->assertXlsxColumnWidthEquals(
            14.43,
            $sheet,
            'D'
        );
    }

    public function testAssertXlsxColumnWidthEqualsShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.C column width does not equal expected value.'
        );

        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxColumnWidthEquals(
            50,
            $sheet,
            'C'
        );
    }

    public function testAssertXlsxSheetRowsCount()
    {
        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxSheetRowsCount(
            3,
            $sheet
        );

        // |A       |B      |
        //1|merged A1:B2    |
        //2|                |
        $sheet = $this->loadSheet('example_merged.xlsx');
        $this->assertXlsxSheetRowsCount(1, $sheet);
    }

    public function testAssertXlsxSheetRowsCountShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Not empty rows count does not equal expected value.'
        );

        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxSheetRowsCount(
            10,
            $sheet
        );
    }

    public function testAssertXlsxSheetColumnsCount()
    {
        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxSheetColumnsCount(4, $sheet);

        // |A       |B      |
        //1|merged A1:B2    |
        //2|                |
        $sheet = $this->loadSheet('example_merged.xlsx');
        $this->assertXlsxSheetColumnsCount(1, $sheet);
    }

    public function testAssertXlsxSheetColumnsCountShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Not empty columns count does not equal expected value.'
        );

        // |A           |B     |                 |D          |
        //1|mixed font  |green |horizontal center|wrap text  |
        //2|regular font|red   |vertical center  |width 14.43|
        //3|            |yellow|width 57.29      |           |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxSheetColumnsCount(3, $sheet);
    }

    public function testAssertXlsxCellMerged()
    {
        // |A       |B      |
        //1|merged A1:B2    |
        //2|                |
        $sheet = $this->loadSheet('example_merged.xlsx');

        $this->assertXlsxCellsMerged($sheet, 'A1:B2');
    }

    public function testAssertXlsxCellMergedShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Cells of C1:D2 range are not merged.'
        );

        // |A       |B      |
        //1|merged A1:B2    |
        //2|                |
        $sheet = $this->loadSheet('example_merged.xlsx');

        $this->assertXlsxCellsMerged($sheet, 'C1:D2');
    }

    public function testAssertXlsxCellMergedPartlyMergedRangeShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Cells of A1:A2 range are not merged.'
        );

        // |A       |B      |
        //1|merged A1:B2    |
        //2|                |
        $sheet = $this->loadSheet('example_merged.xlsx');

        $this->assertXlsxCellsMerged($sheet, 'A1:A2');
    }

    public function testAssertXlsxCellFontItalicItalic()
    {
        // |A      |B               |
        //1|regular|regular + italic|
        //2|bold   |italic + bold   |
        //3|italic |bold italic     |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'A3');
    }

    public function testAssertXlsxCellFontItalicItalicAndBold()
    {
        // |A      |B               |
        //1|regular|regular + italic|
        //2|bold   |italic + bold   |
        //3|italic |bold italic     |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'B3');
    }

    public function testAssertXlsxCellFontItalicStartsWithItalic()
    {
        // |A      |B               |
        //1|regular|regular + italic|
        //2|bold   |italic + bold   |
        //3|italic |bold italic     |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'B2');
    }

    public function testAssertXlsxCellFontItalicNormalThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.A1 cell style is not italic.'
        );

        // |A      |B               |
        //1|regular|regular + italic|
        //2|bold   |italic + bold   |
        //3|italic |bold italic     |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'A1');
    }

    public function testAssertXlsxCellFontItalicBoldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.A2 cell style is not italic.'
        );

        // |A      |B               |
        //1|regular|regular + italic|
        //2|bold   |italic + bold   |
        //3|italic |bold italic     |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'A2');
    }

    public function testAssertXlsxCellFontItalicEndsWithItalicThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.B1 cell style is not italic.'
        );

        // |A      |B               |
        //1|regular|regular + italic|
        //2|bold   |italic + bold   |
        //3|italic |bold italic     |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'B1');
    }

    public function testAssertXlsxCellFontUnderlineUnderline()
    {
        // |A        |B                  |
        //1|regular  |regular + underline|
        //2|bold     |underline + bold   |
        //3|underline|underline italic   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'A3');
    }

    public function testAssertXlsxCellFontUnderlineUnderlineAndItalic()
    {
        // |A        |B                  |
        //1|regular  |regular + underline|
        //2|bold     |underline + bold   |
        //3|underline|underline italic   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'B3');
    }

    public function testAssertXlsxCellFontUnderlineStartsWithUnderline()
    {
        // |A        |B                  |
        //1|regular  |regular + underline|
        //2|bold     |underline + bold   |
        //3|underline|underline italic   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'B2');
    }

    public function testAssertXlsxCellFontUnderlineNormalThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.A1 cell style is not underline.'
        );

        // |A        |B                  |
        //1|regular  |regular + underline|
        //2|bold     |underline + bold   |
        //3|underline|underline italic   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'A1');
    }

    public function testAssertXlsxCellFontUnderlineBoldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.A2 cell style is not underline.'
        );

        // |A        |B                  |
        //1|regular  |regular + underline|
        //2|bold     |underline + bold   |
        //3|underline|underline italic   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'A2');
    }

    public function testAssertXlsxCellFontUnderlineEndsWithUnderlineThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.B1 cell style is not underline.'
        );

        // |A        |B                  |
        //1|regular  |regular + underline|
        //2|bold     |underline + bold   |
        //3|underline|underline italic   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'B1');
    }

    /**
     * Loads the xlsx file content.
     */
    private function loadSheet(string $filename): Worksheet
    {
        $xlsxReader = new XlsxReader();

        $filepath = __DIR__.DIRECTORY_SEPARATOR.'/xlsxFiles/'.$filename;

        $spreadsheet = $xlsxReader->load($filepath);

        return $spreadsheet->getActiveSheet();
    }
}
