<?php

namespace Tests;

use AveSystems\XlsxTestUtils\XlsxAssertsTrait;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PHPUnit\Framework\ExpectationFailedException;
use PHPUnit\Framework\TestCase;
use const DIRECTORY_SEPARATOR;

class XlsxAssertsTraitTest extends TestCase
{
    use XlsxAssertsTrait;

    public function testAssertXlsxCellValueEquals_Successful()
    {
        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellValueEquals('смешанный шрифт', $sheet, 'A1');
        $this->assertXlsxCellValueEquals('нормальный шрифт', $sheet, 'A2');
        $this->assertXlsxCellValueEquals('зелёный', $sheet, 'B1');
        $this->assertXlsxCellValueEquals('красный', $sheet, 'B2');
        $this->assertXlsxCellValueEquals('жёлтый', $sheet, 'B3');
    }

    public function testAssertXlsxCellValueEquals_ShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Значение в ячейке A1 не соответствует ожидаемому.'
        );

        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellValueEquals('смешанный ШРИФТ', $sheet, 'A1');
    }

    public function testAssertXlsxCellEmpty()
    {
        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellEmpty($sheet, 'A3');
    }

    public function testAssertXlsxCellEmpty_ShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Ячейка B1 не пуста.'
        );

        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellEmpty($sheet, 'B1');
    }

    public function testAssertXlsxCellFontColorEquals()
    {
        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
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
            'FF34A853',//зелёный
            $sheet,
            'B1'
        );
        $this->assertXlsxCellFontColorEquals(
            'FFEA4335',//красный
            $sheet,
            'B2'
        );
        $this->assertXlsxCellFontColorEquals(
            'FFFBBC04',//жёлтый
            $sheet,
            'B3'
        );
    }

    public function testAssertXlsxCellFontColorEquals_ShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Цвет текста в ячейке A1 не соответствует ожидаемому.'
        );

        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellFontColorEquals(
            Color::COLOR_GREEN,
            $sheet,
            'A1'
        );
    }

    public function testAssertXlsxCellBackgroundColorEquals()
    {
        // |A      |B      |C      |D      |E     |
        //1|зелёный|зелёный|зелёный|       |      |
        //2|       |жёлтый |жёлтый |жёлтый |жёлтый|
        //3|красный|красный|красный|красный|синий |
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

    public function testAssertXlsxCellBackgroundColorEquals_ShouldThrowException()
    {
        // |A      |B      |C      |D      |E     |
        //1|зелёный|зелёный|зелёный|       |      |
        //2|       |жёлтый |жёлтый |жёлтый |жёлтый|
        //3|красный|красный|красный|красный|синий |
        $sheet = $this->loadSheet('example_background.xlsx');

        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Цвет фона в ячейке "D3" не соответствует ожидаемому.'
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
        // |A      |B      |C      |D      |E     |
        //1|зелёный|зелёный|зелёный|       |      |
        //2|       |жёлтый |жёлтый |жёлтый |жёлтый|
        //3|красный|красный|красный|красный|синий |
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

    public function testAssertXlsxCellsBackgroundColorEquals_ShouldThrowException()
    {
        // |A      |B      |C      |D      |E     |
        //1|зелёный|зелёный|зелёный|       |      |
        //2|       |жёлтый |жёлтый |жёлтый |жёлтый|
        //3|красный|красный|красный|красный|синий |
        $sheet = $this->loadSheet('example_background.xlsx');

        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Цвет фона в ячейке "E3" не соответствует ожидаемому.'
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
        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellHorizontalAlignmentEquals(
            Alignment::HORIZONTAL_CENTER,
            $sheet,
            'C1'
        );
    }

    public function testAssertXlsxCellHorizontalAlignmentEquals_ShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Горизонтальное выравнивание в ячейке C2 '.
            'не соответствует ожидаемому.'
        );

        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellHorizontalAlignmentEquals(
            Alignment::HORIZONTAL_CENTER,
            $sheet,
            'C2'
        );
    }

    public function testAssertXlsxCellVerticalAlignmentEquals()
    {
        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellVerticalAlignmentEquals(
            Alignment::VERTICAL_CENTER,
            $sheet,
            'C2'
        );
    }

    public function testAssertXlsxCellVerticalAlignmentEquals_ShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Вертикальное выравнивание в ячейке C1 '.
            'не соответствует ожидаемому.'
        );

        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellVerticalAlignmentEquals(
            Alignment::VERTICAL_CENTER,
            $sheet,
            'C1'
        );
    }

    public function testAssertXlsxCellWrapTextAlignmentTrue()
    {
        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellWrapTextAlignmentTrue(
            $sheet,
            'D1'
        );
    }

    public function testAssertXlsxCellWrapTextAlignmentTrue_ShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Для ячейки C1 не задано оборачивание текста вокруг её границ.'
        );

        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxCellWrapTextAlignmentTrue(
            $sheet,
            'C1'
        );
    }

    public function testAssertXlsxColumnWidthEquals()
    {
        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
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

    public function testAssertXlsxColumnWidthEquals_ShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Ширина ячейки C не соответствует ожидаемой.'
        );

        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxColumnWidthEquals(
            50,
            $sheet,
            'C'
        );
    }

    public function testAssertXlsxSheetRowsCount()
    {
        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxSheetRowsCount(
            3,
            $sheet
        );
    }

    public function testAssertXlsxSheetRowsCount_ShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Неверное количество строк.'
        );

        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxSheetRowsCount(
            10,
            $sheet
        );
    }

    public function testAssertXlsxSheetColumnsCount()
    {
        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxSheetColumnsCount(4, $sheet);

        // |A          |B      |
        //1|объединены A1:B2   |
        $sheet = $this->loadSheet('example_merged.xlsx');
        $this->assertXlsxSheetColumnsCount(1, $sheet);
    }

    public function testAssertXlsxSheetColumnsCount_ShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Неверное количество столбцов.'
        );

        // |A               |B      |C                           |D                             |
        //1|смешанный шрифт |зелёный|горизонтальное центрирование|обтекание текста вокруг границ|
        //2|нормальный шрифт|красный|вертикальное центрирование  |ширина 14.43                  |
        //3|                |жёлтый |ширина 57.29                |                              |
        $sheet = $this->loadSheet('example.xlsx');

        $this->assertXlsxSheetColumnsCount(3, $sheet);
    }

    public function testAssertXlsxCellMerged()
    {
        // |A          |B      |
        //1|объединены A1:B2   |
        //2|                   |
        $sheet = $this->loadSheet('example_merged.xlsx');

        $this->assertXlsxCellsMerged($sheet, 'A1:B2');
    }

    public function testAssertXlsxCellMerged_ShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Ячейки не объединены в диапазоне C1:D2.'
        );

        // |A          |B      |
        //1|объединены A1:B2   |
        //2|                   |
        $sheet = $this->loadSheet('example_merged.xlsx');

        $this->assertXlsxCellsMerged($sheet, 'C1:D2');
    }

    public function testAssertXlsxCellMerged_PartlyMergedRange_ShouldThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Ячейки не объединены в диапазоне A1:A2.'
        );

        // |A          |B      |
        //1|объединены A1:B2   |
        //2|                   |
        $sheet = $this->loadSheet('example_merged.xlsx');

        $this->assertXlsxCellsMerged($sheet, 'A1:A2');
    }

    public function testAssertXlsxCellFontItalic_Italic()
    {
        // |A      |B               |
        //1|обычный|обычный + курсив|
        //2|жирный |курсив + жирный |
        //3|курсив |жирный курсив   |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'A3');
    }

    public function testAssertXlsxCellFontItalic_ItalicAndBold()
    {
        // |A      |B               |
        //1|обычный|обычный + курсив|
        //2|жирный |курсив + жирный |
        //3|курсив |жирный курсив   |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'B3');
    }

    public function testAssertXlsxCellFontItalic_StartsWithItalic()
    {
        // |A      |B               |
        //1|обычный|обычный + курсив|
        //2|жирный |курсив + жирный |
        //3|курсив |жирный курсив   |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'B2');
    }

    public function testAssertXlsxCellFontItalic_Normal_ThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Текст в ячейке A1 не выделен курсивом.'
        );

        // |A      |B               |
        //1|обычный|обычный + курсив|
        //2|жирный |курсив + жирный |
        //3|курсив |жирный курсив   |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'A1');
    }

    public function testAssertXlsxCellFontItalic_Bold_ThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Текст в ячейке A2 не выделен курсивом.'
        );

        // |A      |B               |
        //1|обычный|обычный + курсив|
        //2|жирный |курсив + жирный |
        //3|курсив |жирный курсив   |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'A2');
    }

    public function testAssertXlsxCellFontItalic_EndsWithItalic_ThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Текст в ячейке B1 не выделен курсивом.'
        );

        // |A      |B               |
        //1|обычный|обычный + курсив|
        //2|жирный |курсив + жирный |
        //3|курсив |жирный курсив   |
        $sheet = $this->loadSheet('example_italic.xlsx');

        $this->assertXlsxCellFontItalic($sheet, 'B1');
    }

    public function testAssertXlsxCellFontUnderline_Underline()
    {
        // |A           |B                     |
        //1|обычный     |обычный + подчёркнутый|
        //2|жирный      |подчёркнутый + жирный |
        //3|подчёркнутый|подчёркнутый курсив   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'A3');
    }

    public function testAssertXlsxCellFontUnderline_UnderlineAndItalic()
    {
        // |A           |B                     |
        //1|обычный     |обычный + подчёркнутый|
        //2|жирный      |подчёркнутый + жирный |
        //3|подчёркнутый|подчёркнутый курсив   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'B3');
    }

    public function testAssertXlsxCellFontUnderline_StartsWithUnderline()
    {
        // |A           |B                     |
        //1|обычный     |обычный + подчёркнутый|
        //2|жирный      |подчёркнутый + жирный |
        //3|подчёркнутый|подчёркнутый курсив   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'B2');
    }

    public function testAssertXlsxCellFontUnderline_Normal_ThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Текст в ячейке A1 не подчёркнут.'
        );

        // |A           |B                     |
        //1|обычный     |обычный + подчёркнутый|
        //2|жирный      |подчёркнутый + жирный |
        //3|подчёркнутый|подчёркнутый курсив   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'A1');
    }

    public function testAssertXlsxCellFontUnderline_Bold_ThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Текст в ячейке A2 не подчёркнут.'
        );

        // |A           |B                     |
        //1|обычный     |обычный + подчёркнутый|
        //2|жирный      |подчёркнутый + жирный |
        //3|подчёркнутый|подчёркнутый курсив   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'A2');
    }

    public function testAssertXlsxCellFontUnderline_EndsWithUnderline_ThrowException()
    {
        $this->expectException(ExpectationFailedException::class);
        $this->expectExceptionMessageMatches(
            '.Текст в ячейке B1 не подчёркнут.'
        );

        // |A           |B                     |
        //1|обычный     |обычный + подчёркнутый|
        //2|жирный      |подчёркнутый + жирный |
        //3|подчёркнутый|подчёркнутый курсив   |
        $sheet = $this->loadSheet('example_underline.xlsx');

        $this->assertXlsxCellFontUnderline($sheet, 'B1');
    }

    /**
     * Загружает содержимое xlsx-файла.
     */
    private function loadSheet(string $filename): Worksheet
    {
        $xlsxReader = new XlsxReader();

        $filepath = __DIR__.DIRECTORY_SEPARATOR.'/xlsxFiles/'.$filename;

        $spreadsheet = $xlsxReader->load($filepath);

        return $spreadsheet->getActiveSheet();
    }
}
