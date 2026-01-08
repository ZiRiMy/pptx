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
     * Add a resource to the presentation.
     *
     * @param ResourceInterface $resource Resource to add
     * @return string|null The resource ID or null
     */
    public function addResource(ResourceInterface $resource): ?string
    {
        if ($resource instanceof NoteMaster) {
            // Check if this NoteMaster is already registered
            $existingRId = $this->findExistingResourceId($resource);
            if ($existingRId !== null) {
                return $existingRId;
            }

            $rId = parent::addResource($resource);
            
            // Create notesMasterIdLst if it doesn't exist
            if (!count($this->content->xpath('p:notesMasterIdLst'))) {
                $this->content->addChild('p:notesMasterIdLst');
            }
            
            // Always add the notesMasterId to the list
            $ref = $this->content->xpath('p:notesMasterIdLst')[0]->addChild('notesMasterId');
            $ref->addAttribute('r:id', $rId, $this->namespaces['r']);

            return $rId;
        }

        if ($resource instanceof Slide) {
            $rId = parent::addResource($resource);

            $currentSlides = $this->content->xpath('p:sldIdLst/p:sldId');

            // PowerPoint slide IDs must be sequential starting from 256
            // Calculate the next sequential ID based on the number of existing slides
            $nextId = 256 + count($currentSlides);

            $ref = $this->content->xpath('p:sldIdLst')[0]->addChild('sldId');
            $ref->addAttribute('id', (string) $nextId);
            $ref->addAttribute('r:id', $rId, $this->namespaces['r']);

            // Add slide to section if it has source section info
            $sourceSection = $resource->getSourceSection();
            if ($sourceSection !== null) {
                $this->addSlideToSection($nextId, $sourceSection['name'], $sourceSection['id']);
            }

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

        return null;
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
