<?php

namespace AveSystems\XlsxTestUtils;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PHPUnit\Framework\ExpectationFailedException;

/**
 * It is designed to simplify asserting xlsx file content.
 * It should be used in test classes inherited from "TestCase".
 *
 * @method assertEquals($expected, $actual, string $message = '', float $delta = 0.0, int $maxDepth = 10, bool $canonicalize = false, bool $ignoreCase = false): void
 * @method assertEmpty($actual, string $message = ''): void
 * @method assertTrue($condition, string $message = ''): void
 */
trait XlsxAssertsTrait
{
    /**
     * Asserts that the cell value and the given value are equal (ignoring
     * style).
     *
     * @param string    $expectedValue
     * @param Worksheet $sheet
     * @param string    $cellCoordinate for example, "A1"
     */
    private function assertXlsxCellValueEquals(
        string $expectedValue,
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $this->assertCellCoordinateIsValid($cellCoordinate, $sheet);

        /**
         * Suppress inspection because assertCellCoordinateIsValid called above
         * checked all conditions.
         *
         * @see SpreadsheetException
         *
         * @noinspection PhpUnhandledExceptionInspection
         */
        $cell = $sheet->getCell($cellCoordinate);
        $realValue = $cell->getValue();
        if ($realValue instanceof RichText) {
            $realValue = $realValue->getPlainText();
        }
        $this->assertEquals(
            $expectedValue,
            $realValue,
            "{$cellCoordinate} cell value does not equal expected value"
        );
    }

    /**
     * Asserts that the cell value is empty.
     *
     * @param Worksheet $sheet
     * @param string    $cellCoordinate for example, "A1"
     */
    private function assertXlsxCellEmpty(
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $this->assertCellCoordinateIsValid($cellCoordinate, $sheet);

        /**
         * Suppress inspection because assertCellCoordinateIsValid called above
         * checked all conditions.
         *
         * @see SpreadsheetException
         *
         * @noinspection PhpUnhandledExceptionInspection
         */
        $cell = $sheet->getCell($cellCoordinate);
        $value = $cell ? $cell->getValue() : $cell;

        $this->assertEmpty(
            $value,
            "{$cellCoordinate} cell value is not empty"
        );
    }

    /**
     * Asserts that the cell font color and the given color are equal.
     *
     * @param string    $expectedArgbColor
     * @param Worksheet $sheet
     * @param string    $cellCoordinate     for example, "A1"
     */
    private function assertXlsxCellFontColorEquals(
        string $expectedArgbColor,
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $this->assertCellCoordinateIsValid($cellCoordinate, $sheet);

        /**
         * Suppress inspection because assertCellCoordinateIsValid called above
         * checked all conditions.
         *
         * @see SpreadsheetException
         *
         * @noinspection PhpUnhandledExceptionInspection
         */
        $cell = $sheet->getCell($cellCoordinate);

        $this->assertEquals(
            $expectedArgbColor,
            $cell->getStyle()->getFont()->getColor()->getARGB(),
            "{$cellCoordinate} cell font color does not equal expected value"
        );
    }

    /**
     * Asserts that the cell font style is italic.
     *
     * @param Worksheet $sheet
     * @param string    $cellCoordinate for example, "A1"
     */
    private function assertXlsxCellFontItalic(
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $this->assertCellCoordinateIsValid($cellCoordinate, $sheet);

        /**
         * Suppress inspection because assertCellCoordinateIsValid called above
         * checked all conditions.
         *
         * @see SpreadsheetException
         *
         * @noinspection PhpUnhandledExceptionInspection
         */
        $cell = $sheet->getCell($cellCoordinate);

        $this->assertTrue(
            $cell->getStyle()->getFont()->getItalic(),
            "{$cellCoordinate} cell style is not italic"
        );
    }

    /**
     * Asserts that the cell font style is underline.
     *
     * @param Worksheet $sheet
     * @param string    $cellCoordinate for example, "A1"
     */
    private function assertXlsxCellFontUnderline(
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $this->assertCellCoordinateIsValid($cellCoordinate, $sheet);

        /**
         * Suppress inspection because assertCellCoordinateIsValid called above
         * checked all conditions.
         *
         * @see SpreadsheetException
         *
         * @noinspection PhpUnhandledExceptionInspection
         */
        $cell = $sheet->getCell($cellCoordinate);

        $this->assertEquals(
            Font::UNDERLINE_SINGLE,
            $cell->getStyle()->getFont()->getUnderline(),
            "{$cellCoordinate} cell style is not underline"
        );
    }

    /**
     * Asserts that the cell horizontal alignment and the given horizontal
     * alignment are equal.
     *
     * @param string    $expectedHorizontalAlignment one of the constants in
     *                                               Alignment class, see "@see"
     * @param Worksheet $sheet
     * @param string    $cellCoordinate              for example, "A1"
     *
     * @see Alignment::HORIZONTAL_CENTER
     */
    private function assertXlsxCellHorizontalAlignmentEquals(
        string $expectedHorizontalAlignment,
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $this->assertCellCoordinateIsValid($cellCoordinate, $sheet);

        /**
         * Suppress inspection because assertCellCoordinateIsValid called above
         * checked all conditions.
         *
         * @see SpreadsheetException
         *
         * @noinspection PhpUnhandledExceptionInspection
         */
        $cell = $sheet->getCell($cellCoordinate);

        $this->assertEquals(
            $expectedHorizontalAlignment,
            $cell->getStyle()->getAlignment()->getHorizontal(),
            "{$cell->getCoordinate()} cell horizontal alignment does not ".
            'equal expected value'
        );
    }

    /**
     * Asserts that the cell vertical alignment and the given vertical
     * alignment are equal.
     *
     * @param string    $expectedVerticalAlignment one of the constants in
     *                                             Alignment class, see "@see"
     * @param Worksheet $sheet
     * @param string    $cellCoordinate for example, "A1"
     *
     * @see Alignment::VERTICAL_CENTER
     */
    private function assertXlsxCellVerticalAlignmentEquals(
        string $expectedVerticalAlignment,
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $this->assertCellCoordinateIsValid($cellCoordinate, $sheet);

        /**
         * Suppress inspection because assertCellCoordinateIsValid called above
         * checked all conditions.
         *
         * @see SpreadsheetException
         *
         * @noinspection PhpUnhandledExceptionInspection
         */
        $cell = $sheet->getCell($cellCoordinate);

        $this->assertEquals(
            $expectedVerticalAlignment,
            $cell->getStyle()->getAlignment()->getVertical(),
            "{$cell->getCoordinate()} cell vertical alignment does not ".
            'equal expected value'
        );
    }

    /**
     * Asserts that the cell has wrap text alignment.
     *
     * @param Worksheet $sheet
     * @param string    $cellCoordinate for example, "A1"
     */
    private function assertXlsxCellWrapTextAlignmentTrue(
        Worksheet $sheet,
        string $cellCoordinate
    ): void {
        $this->assertCellCoordinateIsValid($cellCoordinate, $sheet);

        /**
         * Suppress inspection because assertCellCoordinateIsValid called above
         * checked all conditions.
         *
         * @see SpreadsheetException
         *
         * @noinspection PhpUnhandledExceptionInspection
         */
        $cell = $sheet->getCell($cellCoordinate);

        $this->assertTrue(
            $cell->getStyle()->getAlignment()->getWrapText(),
            "{$cellCoordinate} cell does not have wrap text"
        );
    }

    /**
     * Asserts that the column width and the given width are equal.
     *
     * @param float     $expectedWidth
     * @param Worksheet $sheet
     * @param string    $columnLetter for example, "A"
     */
    private function assertXlsxColumnWidthEquals(
        float $expectedWidth,
        Worksheet $sheet,
        string $columnLetter
    ): void {
        $columnDimension = $sheet->getColumnDimension($columnLetter);
        $this->assertEquals(
            $expectedWidth,
            $columnDimension->getWidth(),
            "$columnLetter column width does not equal expected value"
        );
    }

    /**
     * Asserts that max rows that contain data count equals given value.
     *
     * @param int       $count
     * @param Worksheet $sheet
     */
    private function assertXlsxSheetRowsCount(
        int $count,
        Worksheet $sheet
    ): void {
        $this->assertCount(
            $count,
            $sheet->toArray(),
            'Not empty rows count does not equal expected value'
        );
    }

    /**
     * Asserts that max columns that contain data count equals given value.
     *
     * @param int       $count
     * @param Worksheet $sheet
     */
    private function assertXlsxSheetColumnsCount(
        int $count,
        Worksheet $sheet
    ): void {
        $this->assertCount(
            $count,
            $sheet->toArray()[0],
            'Not empty columns count does not equal expected value'
        );
    }

    /**
     * Asserts that the cells of given range are merged.
     *
     * @param Worksheet $sheet
     * @param string    $cellRange for example, "A1:A2"
     */
    private function assertXlsxCellsMerged(
        Worksheet $sheet,
        string $cellRange
    ): void {
        [$firstCellCoordinate] = explode(':', $cellRange);

        $this->assertCellCoordinateIsValid($firstCellCoordinate, $sheet);

        /**
         * Suppress inspection because assertCellCoordinateIsValid called above
         * checked all conditions.
         *
         * @see SpreadsheetException
         *
         * @noinspection PhpUnhandledExceptionInspection
         */
        $cell = $sheet->getCell($firstCellCoordinate);
        $this->assertEquals(
            $cellRange,
            $cell->getMergeRange(),
            "Cells of $cellRange range are not merged"
        );
    }

    /**
     * Asserts that the cells of given range background colors and the given
     * colors are equal.
     *
     * @param string    $startColor
     * @param string    $endColor
     * @param Worksheet $sheet
     * @param string    $cellRange
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
     * Asserts that the cell background color and the given color are equal.
     *
     * @param string    $startColor
     * @param string    $endColor
     * @param Worksheet $sheet
     * @param string    $cellCoordinate
     */
    private function assertXlsxCellBackgroundColorEquals(
        string $startColor,
        string $endColor,
        Worksheet $sheet,
        string $cellCoordinate
    ) {
        $this->assertCellCoordinateIsValid($cellCoordinate, $sheet);

        /**
         * Suppress inspection because assertCellCoordinateIsValid called above
         * checked all conditions.
         *
         * @see SpreadsheetException
         *
         * @noinspection PhpUnhandledExceptionInspection
         */
        $fill = $sheet->getCell($cellCoordinate)
            ->getStyle()
            ->getFill();

        $this->assertEquals(
            $startColor,
            $fill->getStartColor()->getARGB(),
            "{$cellCoordinate} cell background start color does not equal ".
            'expected value'
        );

        $this->assertEquals(
            $endColor,
            $fill->getEndColor()->getARGB(),
            "{$cellCoordinate} cell background end color does not equal ".
            'expected value'
        );
    }

    /**
     * Assert that the cell coordinate is not absolute and is not a range of
     * cells.
     *
     * @param string    $cellCoordinate
     * @param Worksheet $sheet
     */
    private function assertCellCoordinateIsValid(
        string $cellCoordinate,
        Worksheet $sheet
    ) {
        try {
            $sheet->getCell($cellCoordinate);
        } catch (SpreadsheetException $e) {
            throw new ExpectationFailedException($e->getMessage());
        }
    }
}
