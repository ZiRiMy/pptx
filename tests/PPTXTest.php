<?php

declare(strict_types=1);

namespace Cristal\Presentation\Tests;

use Cristal\Presentation\PPTX;

class PPTXTest extends TestCase
{
    /**
     * Number of slides in the test PowerPoint.
     */
    private const POWERPOINT_SLIDE_COUNT = 1;

    protected PPTX $pptx;

    public function setUp(): void
    {
        parent::setUp();
        $this->pptx = new PPTX(__DIR__ . '/mock/FIN.pptx');
    }

    /**
     * @test
     */
    public function it_loads_all_slides(): void
    {
        $this->assertEquals(
            self::POWERPOINT_SLIDE_COUNT,
            count($this->pptx->getSlides())
        );
    }

    /**
     * @test
     */
    public function it_merges_two_pptx(): void
    {
        $nbSourceSlides = count($this->pptx->getSlides());
        $pptxToAppend = new PPTX(__DIR__ . '/mock/FIN.pptx');

        $this->pptx->addSlides($pptxToAppend->getSlides());
        $this->pptx->saveAs(self::TMP_PATH . '/merge.pptx');

        $mergedPPTX = new PPTX(self::TMP_PATH . '/merge.pptx');

        $this->assertEquals(
            $nbSourceSlides + count($pptxToAppend->getSlides()),
            count($mergedPPTX->getSlides())
        );
    }

    /**
     * @test
     */
    public function it_returns_optimization_stats(): void
    {
        $stats = $this->pptx->getOptimizationStats();

        $this->assertIsArray($stats);
        $this->assertArrayHasKey('original_size', $stats);
        $this->assertArrayHasKey('optimized_size', $stats);
        $this->assertArrayHasKey('cache_stats', $stats);
    }

    /**
     * @test
     */
    public function it_validates_presentation(): void
    {
        $validation = $this->pptx->validate();

        $this->assertIsArray($validation);
        $this->assertArrayHasKey('valid', $validation);
        $this->assertArrayHasKey('slides', $validation);
        $this->assertArrayHasKey('resources', $validation);
    }

    /**
     * @test
     */
    public function it_returns_config(): void
    {
        $config = $this->pptx->getConfig();

        $this->assertInstanceOf(\Cristal\Presentation\Config\OptimizationConfig::class, $config);
        $this->assertFalse($config->isEnabled('image_compression'));
        $this->assertTrue($config->isEnabled('lazy_loading'));
    }

    /**
     * @test
     */
    public function it_returns_image_cache(): void
    {
        $cache = $this->pptx->getImageCache();

        $this->assertInstanceOf(\Cristal\Presentation\Cache\ImageCache::class, $cache);
        $this->assertEquals(0, $cache->count());
    }

    /**
     * @test
     */
    public function it_merges_without_duplicating_masters(): void
    {
        $source = new PPTX(__DIR__ . '/mock/FIN.pptx');
        $toMerge = new PPTX(__DIR__ . '/mock/FIN.pptx');

        $source->addSlides($toMerge->getSlides());
        $source->saveAs(self::TMP_PATH . '/merge_no_dup.pptx');

        // Verify the structure
        $zip = new \ZipArchive();
        $zip->open(self::TMP_PATH . '/merge_no_dup.pptx');

        $presentation = $zip->getFromName('ppt/presentation.xml');
        $xml = simplexml_load_string($presentation);
        $xml->registerXPathNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // Check SlideMasters: should have only 1
        $slideMasters = $xml->xpath('//p:sldMasterIdLst/p:sldMasterId');
        $this->assertCount(
            1,
            $slideMasters,
            'Merged presentation should have only 1 SlideMaster (no duplication)'
        );

        // Check Slides: should have 2
        $slides = $xml->xpath('//p:sldIdLst/p:sldId');
        $this->assertCount(
            2,
            $slides,
            'Merged presentation should have 2 slides'
        );

        // Check NotesMasters: should have only 1
        $notesMasters = $xml->xpath('//p:notesMasterIdLst/p:notesMasterId');
        $this->assertCount(
            1,
            $notesMasters,
            'Merged presentation should have only 1 NotesMaster (no duplication)'
        );

        $zip->close();

        // Verify the merged presentation can be opened without corruption
        $merged = new PPTX(self::TMP_PATH . '/merge_no_dup.pptx');
        $this->assertCount(2, $merged->getSlides());
    }

    /**
     * @test
     * Test that merged file is valid and can be opened by PowerPoint without repair
     */
    public function it_creates_valid_merged_pptx_file(): void
    {
        $source = new PPTX(__DIR__ . '/mock/FIN.pptx');
        $toMerge = new PPTX(__DIR__ . '/mock/FIN.pptx');

        $source->addSlides($toMerge->getSlides());
        $source->saveAs(self::TMP_PATH . '/valid_merge.pptx');

        // Test 1: File should be a valid ZIP
        $this->assertTrue(
            $this->isValidZipFile(self::TMP_PATH . '/valid_merge.pptx'),
            'Merged file is not a valid ZIP archive'
        );

        // Test 2: Essential PPTX files should exist
        $this->assertTrue(
            $this->hasEssentialPptxFiles(self::TMP_PATH . '/valid_merge.pptx'),
            'Merged file is missing essential PPTX files'
        );

        // Test 3: XML files should be well-formed
        $xmlValidation = $this->validateXmlFiles(self::TMP_PATH . '/valid_merge.pptx');
        $this->assertTrue(
            $xmlValidation['valid'],
            'XML validation failed: ' . implode(', ', $xmlValidation['errors'])
        );

        // Test 4: Relationships should be consistent
        $relsValidation = $this->validateRelationships(self::TMP_PATH . '/valid_merge.pptx');
        $this->assertTrue(
            $relsValidation['valid'],
            'Relationships validation failed: ' . implode(', ', $relsValidation['errors'])
        );

        // Test 5: File can be reopened by the library without errors
        $merged = new PPTX(self::TMP_PATH . '/valid_merge.pptx');
        $this->assertCount(2, $merged->getSlides(), 'Merged file should have 2 slides');
    }

    /**
     * Check if file is a valid ZIP archive
     */
    private function isValidZipFile(string $path): bool
    {
        $zip = new \ZipArchive();
        $result = $zip->open($path, \ZipArchive::CHECKCONS);
        
        if ($result === true) {
            $zip->close();
            return true;
        }
        
        return false;
    }

    /**
     * Check if all essential PPTX files exist
     */
    private function hasEssentialPptxFiles(string $path): bool
    {
        $zip = new \ZipArchive();
        $zip->open($path);

        $essentialFiles = [
            '[Content_Types].xml',
            '_rels/.rels',
            'ppt/presentation.xml',
            'ppt/_rels/presentation.xml.rels',
        ];

        foreach ($essentialFiles as $file) {
            if ($zip->locateName($file) === false) {
                $zip->close();
                return false;
            }
        }

        $zip->close();
        return true;
    }

    /**
     * Validate that all XML files are well-formed
     */
    private function validateXmlFiles(string $path): array
    {
        $zip = new \ZipArchive();
        $zip->open($path);
        
        $errors = [];
        $xmlFiles = [];
        
        // Find all XML files
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $filename = $zip->getNameIndex($i);
            if (str_ends_with($filename, '.xml') || str_ends_with($filename, '.rels')) {
                $xmlFiles[] = $filename;
            }
        }

        // Validate each XML file
        foreach ($xmlFiles as $xmlFile) {
            $content = $zip->getFromName($xmlFile);
            if ($content === false) {
                $errors[] = "Cannot read file: $xmlFile";
                continue;
            }

            libxml_use_internal_errors(true);
            $xml = simplexml_load_string($content);
            
            if ($xml === false) {
                $xmlErrors = libxml_get_errors();
                foreach ($xmlErrors as $error) {
                    $errors[] = "$xmlFile: " . trim($error->message);
                }
                libxml_clear_errors();
            }
        }

        $zip->close();

        return [
            'valid' => empty($errors),
            'errors' => $errors,
            'checked_files' => count($xmlFiles),
        ];
    }

    /**
     * Validate that relationships are consistent
     */
    private function validateRelationships(string $path): array
    {
        $zip = new \ZipArchive();
        $zip->open($path);
        
        $errors = [];
        
        // Check presentation.xml.rels
        $relsContent = $zip->getFromName('ppt/_rels/presentation.xml.rels');
        if ($relsContent === false) {
            $errors[] = 'Missing ppt/_rels/presentation.xml.rels';
        } else {
            $relsXml = simplexml_load_string($relsContent);
            $relsXml->registerXPathNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships');
            
            // Check each relationship points to an existing file
            $relationships = $relsXml->xpath('//r:Relationship');
            foreach ($relationships as $rel) {
                $target = (string)$rel['Target'];
                $type = (string)$rel['Type'];
                
                // Skip external relationships
                if (isset($rel['TargetMode']) && (string)$rel['TargetMode'] === 'External') {
                    continue;
                }
                
                // Construct full path
                $fullPath = 'ppt/' . $target;
                
                // Check if target file exists
                if ($zip->locateName($fullPath) === false) {
                    $errors[] = "Relationship target not found: $fullPath (from presentation.xml.rels)";
                }
            }
        }
        
        // Check slide relationships
        $presentationContent = $zip->getFromName('ppt/presentation.xml');
        if ($presentationContent !== false) {
            $presXml = simplexml_load_string($presentationContent);
            $presXml->registerXPathNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
            $presXml->registerXPathNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
            
            // Get all slide IDs
            $slides = $presXml->xpath('//p:sldId');
            foreach ($slides as $slide) {
                $rId = (string)$slide->attributes('http://schemas.openxmlformats.org/officeDocument/2006/relationships')->id;
                
                // Find the target in rels
                $relsXml = simplexml_load_string($relsContent);
                $relsXml->registerXPathNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships');
                $targetRel = $relsXml->xpath("//r:Relationship[@Id='$rId']");
                
                if (empty($targetRel)) {
                    $errors[] = "Slide relationship not found in rels: $rId";
                }
            }
        }

        $zip->close();

        return [
            'valid' => empty($errors),
            'errors' => $errors,
        ];
    }
}
