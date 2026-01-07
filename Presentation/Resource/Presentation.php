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
            if (!count($this->content->xpath('p:notesMasterIdLst'))) {
                $this->content->addChild('p:notesMasterIdLst');

                $ref = $this->content->xpath('p:notesMasterIdLst')[0]->addChild('notesMasterId');
                $ref->addAttribute('r:id', $rId, $this->namespaces['r']);
            }

            return $rId;
        }

        if ($resource instanceof Slide) {
            $rId = parent::addResource($resource);

            $currentSlides = $this->content->xpath('p:sldIdLst/p:sldId');

            $ref = $this->content->xpath('p:sldIdLst')[0]->addChild('sldId');
            $ref->addAttribute('id', (string) ((int) end($currentSlides)['id'] + 1));
            $ref->addAttribute('r:id', $rId, $this->namespaces['r']);

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
     * Find if a resource is already registered in this presentation.
     * Used to avoid duplicating structural resources (Masters, Themes, etc.).
     *
     * @param ResourceInterface $resource The resource to check
     * @return string|null The existing resource ID if found, null otherwise
     */
    private function findExistingResourceId(ResourceInterface $resource): ?string
    {
        $this->mapResources();
        
        foreach ($this->resources as $rId => $existingResource) {
            if ($existingResource === $resource) {
                return $rId;
            }
        }
        
        return null;
    }
}
