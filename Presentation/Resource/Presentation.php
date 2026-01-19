<?php

declare(strict_types=1);

namespace Cristal\Presentation\Resource;

use Cristal\Presentation\ResourceInterface;

/**
 * Presentation resource class for handling the main presentation.xml file.
 */
class Presentation extends XmlResource
{
    /**
     * Track if sections have been removed to avoid repeated removal
     */
    private static bool $sectionsRemoved = false;

    /**
     * Add a resource to the presentation.
     *
     * @param ResourceInterface $resource Resource to add
     * @return string|null The resource ID or null
     */
    public function addResource(ResourceInterface $resource): ?string
    {
        // Remove sections on first slide addition to avoid incorrect ordering
        if ($resource instanceof Slide && !self::$sectionsRemoved) {
            $this->removeSections();
            self::$sectionsRemoved = true;
        }

        if ($resource instanceof NoteMaster) {
            // PowerPoint only supports ONE NoteMaster per presentation
            // Check if a notesMaster already exists and reuse it
            $existing = $this->content->xpath('p:notesMasterIdLst/p:notesMasterId/@r:id');
            if (!empty($existing)) {
                // NoteMaster already exists, return its rId
                return (string)$existing[0];
            }

            // No NoteMaster exists yet, add this one
            $rId = parent::addResource($resource);

            // Create notesMasterIdLst if it doesn't exist
            if (!count($this->content->xpath('p:notesMasterIdLst'))) {
                $this->content->addChild('p:notesMasterIdLst');
            }

            // Add the notesMasterId to the list
            $ref = $this->content->xpath('p:notesMasterIdLst')[0]->addChild('notesMasterId');
            $ref->addAttribute('r:id', $rId, $this->namespaces['r']);

            return $rId;
        }

        if ($resource instanceof Slide) {
            // CRITICAL: Slides must have consecutive rIds (rId2, rId3, rId4...)
            // Find the next available rId specifically for slides
            $rId = $this->getNextSlideRId();

            // Manually add to resources array (bypass parent::addResource which uses max+1)
            $this->resources[$rId] = $resource;

            $currentSlides = $this->content->xpath('p:sldIdLst/p:sldId');

            // PowerPoint slide IDs must be unique and >= 256
            // Find the maximum existing ID and increment it
            $maxId = 255; // Minimum value is 256
            foreach ($currentSlides as $slide) {
                $existingId = (int)$slide['id'];
                if ($existingId > $maxId) {
                    $maxId = $existingId;
                }
            }
            $nextId = $maxId + 1;

            $ref = $this->content->xpath('p:sldIdLst')[0]->addChild('sldId');
            $ref->addAttribute('id', (string) $nextId);
            $ref->addAttribute('r:id', $rId, $this->namespaces['r']);

            // DISABLED: Section copying during merge causes incorrect slide order in sections
            // Sections are organizational features that should be recreated manually after merge
            // The problem is that slides are added to sections in processing order, not final order
            // TODO: Implement proper section reconstruction after all slides are added
            // $sourceSection = $resource->getSourceSection();
            // if ($sourceSection !== null) {
            //     $this->addSlideToSection($nextId, $sourceSection['name'], $sourceSection['id']);
            // }

            return $rId;
        }

        if ($resource instanceof SlideMaster) {
            // Check if this SlideMaster is already registered
            $existingRId = $this->findExistingResourceId($resource);
            if ($existingRId !== null) {
                return $existingRId;
            }

            $rId = parent::addResource($resource);

            $ref = $this->content->xpath('p:sldMasterIdLst')[0]->addChild('sldMasterId');
            $ref->addAttribute('id', (string) self::getUniqueID());
            $ref->addAttribute('r:id', $rId, $this->namespaces['r']);

            return $rId;
        }

        if ($resource instanceof HandoutMaster) {
            $rId = parent::addResource($resource);
            $ref = $this->content->xpath('p:handoutMasterIdLst')[0]->addChild('handoutMasterId');
            $ref->addAttribute('r:id', $rId, $this->namespaces['r']);

            return $rId;
        }

        // For all other resources (presProps, viewProps, theme, tableStyles, etc.),
        // use the parent's generic addResource() to register them in the .rels file
        return parent::addResource($resource);
    }

    /**
     * Remove all sections from the presentation.
     * Sections become invalid when merging presentations, so they should be removed.
     */
    protected function removeSections(): void
    {
        // Register p14 namespace
        $this->content->registerXPathNamespace('p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');

        // Find the ext element containing sectionLst
        $extElements = $this->content->xpath('//p:ext[@uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"]');

        if (!empty($extElements)) {
            // Remove this ext element (contains sections)
            $dom = dom_import_simplexml($extElements[0]);
            $dom->parentNode->removeChild($dom);
        }
    }

    /**
     * Add a slide to a section in the sectionLst.
     * Creates the section if it doesn't exist.
     *
     * @param int $slideId The slide ID to add
     * @param string $sectionName The section name
     * @param string $sectionGuid The section GUID
     */
    protected function addSlideToSection(int $slideId, string $sectionName, string $sectionGuid): void
    {
        // Register p14 namespace
        $this->content->registerXPathNamespace('p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
        
        // Find existing sectionLst in extLst
        $sectionLst = $this->content->xpath('//p14:sectionLst');
        
        if (empty($sectionLst)) {
            // No sections exist yet - we need to create extLst and sectionLst
            // This is complex - for now we'll just skip if no sections exist
            return;
        }
        
        $sectionLst = $sectionLst[0];
        
        // Find section by name
        $existingSection = $sectionLst->xpath("p14:section[@name='$sectionName']");
        
        if (!empty($existingSection)) {
            // Section exists - add slide ID to it
            $section = $existingSection[0];
        } else {
            // Create new section
            $section = $sectionLst->addChild('section', null, 'http://schemas.microsoft.com/office/powerpoint/2010/main');
            $section->addAttribute('name', $sectionName);
            $section->addAttribute('id', $sectionGuid);
            
            // Add sldIdLst to section
            $section->addChild('sldIdLst', null, 'http://schemas.microsoft.com/office/powerpoint/2010/main');
        }
        
        // Add slide ID to section's sldIdLst
        $sldIdLst = $section->xpath('p14:sldIdLst');
        if (!empty($sldIdLst)) {
            $sldId = $sldIdLst[0]->addChild('sldId', null, 'http://schemas.microsoft.com/office/powerpoint/2010/main');
            $sldId->addAttribute('id', (string) $slideId);
        }
    }

    /**
     * Get the next available rId for a slide.
     * Slides should have consecutive rIds, but must not collide with other resources.
     *
     * @return string The next available rId for a slide (e.g., 'rId2', 'rId3', 'rId4'...)
     */
    private function getNextSlideRId(): string
    {
        $this->mapResources();

        // Get all existing rIds (slides and non-slides)
        $allUsedIds = [];
        foreach ($this->resources as $rId => $resource) {
            $allUsedIds[] = (int)str_replace('rId', '', $rId);
        }

        // Find all existing slide rIds
        $slideRIds = [];
        foreach ($this->resources as $rId => $resource) {
            if ($resource instanceof Slide) {
                $slideRIds[] = (int)str_replace('rId', '', $rId);
            }
        }

        // Start from rId2 (rId1 is usually slideMaster)
        $nextId = 2;

        // Try to find the next consecutive rId for slides
        // but skip any rId that's already in use by ANY resource
        sort($slideRIds);
        foreach ($slideRIds as $existingId) {
            if ($existingId == $nextId && !in_array($nextId, $allUsedIds, true)) {
                $nextId++;
            } else if (in_array($nextId, $allUsedIds, true)) {
                // This rId is used by another resource, skip it
                $nextId++;
            } else {
                // Found a gap in slide sequence
                break;
            }
        }

        // Final check: ensure the proposed rId is not in use
        while (in_array($nextId, $allUsedIds, true)) {
            $nextId++;
        }

        return 'rId' . $nextId;
    }

    /**
     * Find if a resource is already registered in this presentation.
     * Used to avoid duplicating structural resources (Masters, Themes, etc.).
     *
     * @param ResourceInterface $resource The resource to check
     * @return string|null The existing resource ID if found, null otherwise
     */
    private function findExistingResourceId(ResourceInterface $resource): ?string
    {
        $this->mapResources();

        // For GenericResource, compare by target path to detect reused resources
        if ($resource instanceof GenericResource) {
            foreach ($this->resources as $rId => $existingResource) {
                if ($existingResource instanceof GenericResource &&
                    $existingResource->getTarget() === $resource->getTarget()) {
                    return $rId;
                }
            }
        }

        // For other resources, compare by reference (original behavior)
        foreach ($this->resources as $rId => $existingResource) {
            if ($existingResource === $resource) {
                return $rId;
            }
        }

        return null;
    }
}
