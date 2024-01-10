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

namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Style\TOC as TOCStyle;
use PhpOffice\PhpWord\Style;

/**
 * Table of contents.
 */
class TOC extends AbstractElement
{
    /**
     * TOC style.
     *
     * @var \PhpOffice\PhpWord\Style\TOC
     */
    private $tocStyle;

    /**
     * Font style.
     *
     * @var \PhpOffice\PhpWord\Style\Font|string
     */
    private $fontStyle;

    /**
     * Min title depth to show.
     *
     * @var int
     */
    private $minDepth = 1;

    /**
     * Max title depth to show.
     *
     * @var int
     */
    private $maxDepth = 9;

    /**
     * Set of styles corresponding to each depth level.
     *
     * @var array
     */
    private $styles = [];

    /**
     * Use numbering prefix in titles.
     *
     * @var bool
     */
    private $useNumbering = true;

    /**
     * Create a new Table-of-Contents Element.
     *
     * @param mixed $styles
     * @param mixed $fontStyle
     * @param array $tocStyle
     * @param int $minDepth
     * @param int $maxDepth
     */
    public function __construct($styles = null, $fontStyle = null, $tocStyle = null, $useNumberingPrefix = false, $minDepth = 1, $maxDepth = 9)
    {
        $this->tocStyle = new TOCStyle();

        if (null !== $tocStyle && is_array($tocStyle)) {
            $this->tocStyle->setStyleByArray($tocStyle);
        }

        if (null !== $fontStyle && is_array($fontStyle)) {
            $this->fontStyle = new Font();
            $this->fontStyle->setStyleByArray($fontStyle);
        } else {
            $this->fontStyle = $fontStyle;
        }

        $this->minDepth = $minDepth;
        $this->maxDepth = $maxDepth;

        $this->styles = $styles;
        $this->useNumbering = $useNumberingPrefix;

        foreach ($this->styles as $style) {
            if (!is_string($style)) {
                throw new \InvalidArgumentException('Style must be a string');
            }

            $s = Style::getStyle($style);

            if($s === null) {
                throw new \InvalidArgumentException('Style "'.$style.'" not found');
            }

            if (!($s instanceof Style\Font) && !($s instanceof Style\Paragraph)) {
                throw new \InvalidArgumentException('Style "'.$style.'" must be a font or a paragprah style');
            }
        }
    }

    /**
     * Get all titles.
     *
     * @return array
     */
    public function getTitles()
    {
        if (!$this->phpWord instanceof PhpWord) {
            return [];
        }

        $titles = $this->phpWord->getTitles()->getItems();
        foreach ($titles as $i => $title) {
            /** @var \PhpOffice\PhpWord\Element\Title $title Type hint */
            $depth = $title->getDepth();
            if ($this->minDepth > $depth) {
                unset($titles[$i]);
            }
            if (($this->maxDepth != 0) && ($this->maxDepth < $depth)) {
                unset($titles[$i]);
            }
        }

        return $titles;
    }

    /**
     * Get title styles
     *
     * @return ?array
     */
    public function getTitleStyles()
    {
        return $this->styles;
    }

    /**
     * Get TOC Style.
     *
     * @return \PhpOffice\PhpWord\Style\TOC
     */
    public function getStyleTOC()
    {
        return $this->tocStyle;
    }

    /**
     * Get Font Style.
     *
     * @return \PhpOffice\PhpWord\Style\Font|string
     */
    public function getStyleFont()
    {
        return $this->fontStyle;
    }

    /**
     * Set max depth.
     *
     * @param int $value
     */
    public function setMaxDepth($value): void
    {
        $this->maxDepth = $value;
    }

    /**
     * Get Max Depth.
     *
     * @return int Max depth of titles
     */
    public function getMaxDepth()
    {
        return $this->maxDepth;
    }

    /**
     * Set min depth.
     *
     * @param int $value
     */
    public function setMinDepth($value): void
    {
        $this->minDepth = $value;
    }

    /**
     * Get Min Depth.
     *
     * @return int Min depth of titles
     */
    public function getMinDepth()
    {
        return $this->minDepth;
    }

    /**
     * Set use numbering.
     *
     * @param bool $value
     */
    public function setUseNumbering($value): void
    {
        $this->useNumbering = $value;
    }

    /**
     * Get use numbering.
     *
     * @return bool
     */
    public function getUseNumbering()
    {
        return $this->useNumbering;
    }
}
