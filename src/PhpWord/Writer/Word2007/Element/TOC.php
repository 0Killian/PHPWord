<?php

/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @see         https://github.com/PHPOffice/PHPWord
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Writer\Word2007\Element;

use PhpOffice\PhpWord\Element\Title;
use PhpOffice\PhpWord\Element\TOC as TOCElement;
use PhpOffice\PhpWord\Shared\XMLWriter;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Writer\Word2007\Style\Font as FontStyleWriter;
use PhpOffice\PhpWord\Writer\Word2007\Style\Paragraph as ParagraphStyleWriter;
use PhpOffice\PhpWord\Style;

/**
 * TOC element writer.
 *
 * @since 0.10.0
 */
class TOC extends AbstractElement
{
    /**
     * Write element.
     */
    public function write(): void
    {
        $xmlWriter = $this->getXmlWriter();
        $element = $this->getElement();
        if (!$element instanceof TOCElement) {
            return;
        }

        $titles = $element->getTitles();
        $writeFieldMark = true;

        $indices = [];

        foreach ($titles as $title) {
            $depth = $title->getDepth();
            if (count($indices) < $depth) {
                $indices[] = 0;
            } else {
                $indices = array_slice($indices, 0, $depth);
                $indices[$depth - 1]++;
            }
            $this->writeTitle($xmlWriter, $element, $title, $writeFieldMark, $indices);
            if ($writeFieldMark) {
                $writeFieldMark = false;
            }
        }

        $xmlWriter->startElement('w:p');
        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'end');
        $xmlWriter->endElement();
        $xmlWriter->endElement();
        $xmlWriter->endElement();
    }

    /**
     * Write title.
     */
    private function writeTitle(XMLWriter $xmlWriter, TOCElement $element, Title $title, bool $writeFieldMark, array $indices): void
    {
        $tocStyle = $element->getStyleTOC();
        $fontStyle = $element->getStyleFont();
        $isObject = ($fontStyle instanceof Font) ? true : false;
        $rId = $title->getRelationId();
        $indent = (int) (($title->getDepth() - 1) * $tocStyle->getIndent());

        $styles = $element->getTitleStyles();
        if (null !== $styles && isset($styles[$title->getDepth() - 1])) {
            $style = $styles[$title->getDepth() - 1];

            if (is_string($style)) {
                $style = Style::getStyle($style);
            }

            if ($style instanceof Font) {
                $fontStyle = $style;
                $isObject = true;
            }
        }

        if (is_string($fontStyle)) {
            $fontStyle = Style::getStyle($fontStyle);
            $isObject = true;
        }

        $titleFontStyle = $title->getStyle();

        if (is_string($titleFontStyle)) {
            $titleFontStyle = Style::getStyle($titleFontStyle);
        }

        if (null !== $titleFontStyle) {
            $titleParagraphStyle = $titleFontStyle->getParagraph();
        } else {
            $titleParagraphStyle = null;
        }

        $xmlWriter->startElement('w:p');

        // Write style and field mark
        $this->writeStyle($xmlWriter, $element, $indent, $title->getDepth());
        if ($writeFieldMark) {
            $this->writeFieldMark($xmlWriter, $element);
        }

        // Hyperlink
        $xmlWriter->startElement('w:hyperlink');
        $xmlWriter->writeAttribute('w:anchor', "_Toc{$rId}");
        $xmlWriter->writeAttribute('w:history', '1');

        // Title text
        $xmlWriter->startElement('w:r');
        if ($isObject) {
            $styleWriter = new FontStyleWriter($xmlWriter, $fontStyle);
            $styleWriter->write();
        }
        $xmlWriter->startElement('w:t');

        $titleText = $title->getText();

        if ($element->getUseNumbering() && $titleParagraphStyle !== null && $titleParagraphStyle->getNumStyle() !== null) {
            $numStyle = $titleParagraphStyle->getNumStyle();

            if (is_string($numStyle)) {
                $numStyle = Style::getStyle($numStyle);
            }

            if ($numStyle->getType() === 'multilevel' || $numStyle->getType() === 'hybridMultilevel') {
                $levels = $numStyle->getLevels();

                foreach($levels as $i => $level) {
                    if (isset($indices[$i])) {
                        $indices[$i] += $level->getStart();
                    }
                }

                if (isset($levels[$title->getDepth() - 1])) {
                    $level = $levels[$title->getDepth() - 1];
                    $numberingTextFormat = $level->getText();
                    foreach ($indices as $depth => $index) {
                        $d = $depth + 1;
                        $numberingTextFormat = str_replace("%{$d}", $index, $numberingTextFormat);
                    }

                    $titleText = $numberingTextFormat . " {$titleText}";
                } else {
                    throw new \Exception('Invalid multilevel numbering style');
                }
            }
        }

        $this->writeText(is_string($titleText) ? $titleText : '');

        $xmlWriter->endElement(); // w:t
        $xmlWriter->endElement(); // w:r

        $xmlWriter->startElement('w:r');
        $xmlWriter->writeElement('w:tab', null);
        $xmlWriter->endElement();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'begin');
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:instrText');
        $xmlWriter->writeAttribute('xml:space', 'preserve');
        $xmlWriter->text("PAGEREF _Toc{$rId} \\h");
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        if ($title->getPageNumber() !== null) {
            $xmlWriter->startElement('w:r');
            $xmlWriter->startElement('w:fldChar');
            $xmlWriter->writeAttribute('w:fldCharType', 'separate');
            $xmlWriter->endElement();
            $xmlWriter->endElement();

            $xmlWriter->startElement('w:r');
            $xmlWriter->startElement('w:t');
            $xmlWriter->text((string) $title->getPageNumber());
            $xmlWriter->endElement();
            $xmlWriter->endElement();
        }

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'end');
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $xmlWriter->endElement(); // w:hyperlink

        $xmlWriter->endElement(); // w:p
    }

    /**
     * Write style.
     */
    private function writeStyle(XMLWriter $xmlWriter, TOCElement $element, int $indent, int $depth): void
    {
        $tocStyle = $element->getStyleTOC();
        $fontStyle = $element->getStyleFont();
        $isObject = ($fontStyle instanceof Font) ? true : false;

        $styles = $element->getTitleStyles();
        if (null !== $styles && isset($styles[$depth - 1])) {
            $style = $styles[$depth - 1];

            if (is_string($style)) {
                $style = Style::getStyle($style);
            }

            if ($style instanceof Font) {
                $fontStyle = $style;
                $isObject = true;
            } else {
                $pStyle = clone $style;
            }
        }

        // Paragraph
        if ($isObject && null !== $fontStyle->getParagraph()) {
            $pStyle = clone $fontStyle->getParagraph();
        } else if (!$isObject && is_string($fontStyle)) {
            $f = Style::getStyle($fontStyle);
            if (null !== $f && null !== $f->getParagraph()) {
                $pStyle = clone $f->getParagraph();
            }
        }

        if (isset($pStyle)) {
            if (count($pStyle->getTabs()) === 0) {
                $pStyle->setTabs($tocStyle);
            }

            if ($indent > 0 && $pStyle->getIndent() === null) {
                $pStyle->setIndent($indent);
            }

            $styleWriter = new ParagraphStyleWriter($xmlWriter, $pStyle);
            $styleWriter->write();
        }

        // Font
        if (!empty($fontStyle) && !$isObject) {
            $xmlWriter->startElement('w:rPr');
            $xmlWriter->startElement('w:rStyle');
            $xmlWriter->writeAttribute('w:val', $fontStyle);
            $xmlWriter->endElement();
            $xmlWriter->endElement(); // w:rPr
        }
    }

    /**
     * Write TOC Field.
     */
    private function writeFieldMark(XMLWriter $xmlWriter, TOCElement $element): void
    {
        $minDepth = $element->getMinDepth();
        $maxDepth = $element->getMaxDepth();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'begin');
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:instrText');
        $xmlWriter->writeAttribute('xml:space', 'preserve');
        $xmlWriter->text("TOC \\o {$minDepth}-{$maxDepth} \\h \\z \\u");
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'separate');
        $xmlWriter->endElement();
        $xmlWriter->endElement();
    }
}
