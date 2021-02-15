<?php

namespace AveSystems\XlsxTestUtils;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Нужен для упрощения проверки содержимого в xlsx.
 * Должен использоваться в тестах, наследованных от "TestCase".
 *
 * @method assertEquals($expected, $actual, string $message = '', float $delta = 0.0, int $maxDepth = 10, bool $canonicalize = false, bool $ignoreCase = false): void
 * @method assertEmpty($actual, string $message = ''): void
 * @method assertTrue($condition, string $message = ''): void
 */
trait XlsxAssertsTrait
{
    /**
     * Проверяет, что значение в ячейке соответствует ожидаемому
     * без учёта форматирования.
     *
     * @param string    $expectedValue
     * @param Worksheet $sheet          excel-лист
     * @param string    $cellCoordinate координата ячейки, например, A1
     *
     * @throws Exception
     */
    private function assertXlsxCellValueEquals(
        string $expectedValue,
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $cell = $sheet->getCell($cellCoordinate);
        $realValue = $cell->getValue();
        if ($realValue instanceof RichText) {
            $realValue = $realValue->getPlainText();
        }
        $this->assertEquals(
            $expectedValue,
            $realValue,
            "Значение в ячейке {$cellCoordinate} не соответствует ожидаемому"
        );
    }

    /**
     * Проверяет, что ячейка пустая или значение в ячейке равно пустому значению.
     *
     * @param Worksheet $sheet          excel-лист
     * @param string    $cellCoordinate координата ячейки, например, A1
     *
     * @throws Exception
     */
    private function assertXlsxCellEmpty(
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $cell = $sheet->getCell($cellCoordinate);
        $value = $cell ? $cell->getValue() : $cell;
        $this->assertEmpty(
            $value,
            "Ячейка {$cellCoordinate} не пуста"
        );
    }

    /**
     * Проверяет, что цвет текста в ячейке соответствует ожидаемому.
     *
     * @param string    $expectedArgbColor
     * @param Worksheet $sheet              excel-лист
     * @param string    $cellCoordinate     координата ячейки, например, A1
     *
     * @throws Exception
     */
    private function assertXlsxCellFontColorEquals(
        string $expectedArgbColor,
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $cell = $sheet->getCell($cellCoordinate);
        $this->assertEquals(
            $expectedArgbColor,
            $cell->getStyle()->getFont()->getColor()->getARGB(),
            "Цвет текста в ячейке {$cellCoordinate} не соответствует ".
            'ожидаемому'
        );
    }

    /**
     * Проверяет, что текст в ячейке выделен курсивом.
     *
     * @param Worksheet $sheet          excel-лист
     * @param string    $cellCoordinate координата ячейки, например, A1
     *
     * @throws Exception
     */
    private function assertXlsxCellFontItalic(
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $cell = $sheet->getCell($cellCoordinate);

        $this->assertTrue(
            $cell->getStyle()->getFont()->getItalic(),
            "Текст в ячейке {$cellCoordinate} не выделен курсивом"
        );
    }

    /**
     * Проверяет, что текст в ячейке подчёркнут.
     *
     * @param Worksheet $sheet          excel-лист
     * @param string    $cellCoordinate координата ячейки, например, A1
     *
     * @throws Exception
     */
    private function assertXlsxCellFontUnderline(
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $cell = $sheet->getCell($cellCoordinate);

        $this->assertEquals(
            Font::UNDERLINE_SINGLE,
            $cell->getStyle()->getFont()->getUnderline(),
            "Текст в ячейке {$cellCoordinate} не подчёркнут"
        );
    }

    /**
     * Проверяет, что горизонтальное выравнивание в ячейке
     * соответствует ожидаемому.
     *
     * @param string    $expectedHorizontalAlignment
     * @param Worksheet $sheet          excel-лист
     * @param string    $cellCoordinate координата ячейки, например, A1
     *
     * @throws Exception
     */
    private function assertXlsxCellHorizontalAlignmentEquals(
        string $expectedHorizontalAlignment,
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $cell = $sheet->getCell($cellCoordinate);
        $this->assertEquals(
            $expectedHorizontalAlignment,
            $cell->getStyle()->getAlignment()->getHorizontal(),
            "Горизонтальное выравнивание в ячейке {$cell->getCoordinate()} ".
            'не соответствует ожидаемому'
        );
    }

    /**
     * Проверяет, что вертикальное выравнивание в ячейке
     * соответствует ожидаемому.
     *
     * @param string    $expectedVerticalAlignment
     * @param Worksheet $sheet          excel-лист
     * @param string    $cellCoordinate координата ячейки, например, A1
     *
     * @throws Exception
     */
    private function assertXlsxCellVerticalAlignmentEquals(
        string $expectedVerticalAlignment,
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $cell = $sheet->getCell($cellCoordinate);
        $this->assertEquals(
            $expectedVerticalAlignment,
            $cell->getStyle()->getAlignment()->getVertical(),
            "Вертикальное выравнивание в ячейке {$cellCoordinate} ".
            'не соответствует ожидаемому'
        );
    }

    /**
     * Проверяет, что для ячейки задано оборачивание текста вокруг её границ,
     * чтобы текст не заходил за пределы границ.
     *
     * @param Worksheet $sheet          excel-лист
     * @param string    $cellCoordinate координата ячейки, например, A1
     *
     * @throws Exception
     */
    private function assertXlsxCellWrapTextAlignmentTrue(
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $cell = $sheet->getCell($cellCoordinate);
        $this->assertTrue(
            $cell->getStyle()->getAlignment()->getWrapText(),
            "Для ячейки {$cellCoordinate} не задано оборачивание текста ".
            'вокруг её границ'
        );
    }

    /**
     * Проверяет, что ширина столбца соответствует ожидаемой.
     *
     * @param float     $expectedWidth
     * @param Worksheet $sheet           excel-лист
     * @param string    $cellColumnIndex строковый индекс столбца ячейки, например 'A'
     */
    private function assertXlsxColumnWidthEquals(
        float $expectedWidth,
        Worksheet $sheet,
        string $cellColumnIndex
    ): void {
        $columnDimension = $sheet->getColumnDimension($cellColumnIndex);
        $this->assertEquals(
            $expectedWidth,
            $columnDimension->getWidth(),
            "Ширина ячейки $cellColumnIndex ".
            'не соответствует ожидаемой'
        );
    }

    /**
     * Проверяет, количество непустых строк в excel-листе.
     *
     * @param int       $count
     * @param Worksheet $sheet excel-лист
     */
    private function assertXlsxSheetRowsCount(
        int $count,
        Worksheet $sheet
    ): void {
        $this->assertCount(
            $count,
            $sheet->toArray(),
            'Неверное количество строк'
        );
    }

    /**
     * Проверяет максимальное количество непустых столбцов в excel-листе.
     *
     * @param int       $count
     * @param Worksheet $sheet excel-лист
     */
    private function assertXlsxSheetColumnsCount(
        int $count,
        Worksheet $sheet
    ): void {
        $this->assertCount(
            $count,
            $sheet->toArray()[0],
            'Неверное количество столбцов'
        );
    }

    /**
     * Проверяет, объединены ли ячейки в определенном диапазоне.
     *
     * @param Worksheet $sheet
     * @param string    $cellRange диапазон ячеек, например A1:A2
     *
     * @throws Exception
     */
    private function assertXlsxCellsMerged(
        Worksheet $sheet,
        string $cellRange
    ): void {
        [$firstCellCoordinate] = explode(':', $cellRange);
        $cell = $sheet->getCell($firstCellCoordinate);
        $this->assertEquals(
            $cellRange,
            $cell->getMergeRange(),
            "Ячейки не объединены в диапазоне $cellRange"
        );
    }

    /**
     * Проверяет, что все ячейки в указанном диапазоне имеют указанный фоновый
     * цвет.
     *
     * @param string    $startColor
     * @param string    $endColor
     * @param Worksheet $sheet
     * @param string    $cellRange
     *
     * @throws Exception
     */
    private function assertXlsxCellsBackgroundColorEquals(
        string $startColor,
        string $endColor,
        Worksheet $sheet,
        string $cellRange
    ) {
        [$rangeStart, $rangeEnd] = Coordinate::rangeBoundaries($cellRange);

        for ($col = $rangeStart[0]; $col <= $rangeEnd[0]; ++$col) {
            for ($row = $rangeStart[1]; $row <= $rangeEnd[1]; ++$row) {
                $this->assertXlsxCellBackgroundColorEquals(
                    $startColor,
                    $endColor,
                    $sheet,
                    Coordinate::stringFromColumnIndex($col).$row
                );
            }
        }
    }

    /**
     * Проверяет, что ячейка имеет указанный фоновый цвет.
     *
     * @param string    $startColor
     * @param string    $endColor
     * @param Worksheet $sheet
     *
     * @param string    $cellCoordinate
     *
     * @throws Exception
     */
    private function assertXlsxCellBackgroundColorEquals(
        string $startColor,
        string $endColor,
        Worksheet $sheet,
        string $cellCoordinate
    ) {
        $fill = $sheet->getCell($cellCoordinate)
            ->getStyle()
            ->getFill();

        $this->assertEquals(
            $startColor,
            $fill->getStartColor()->getARGB(),
            sprintf(
                'Цвет фона в ячейке "%s" не соответствует ожидаемому',
                $cellCoordinate
            )
        );

        $this->assertEquals(
            $endColor,
            $fill->getEndColor()->getARGB(),
            sprintf(
                'Цвет фона в ячейке "%s" не соответствует ожидаемому',
                $cellCoordinate
            )
        );
    }
}
