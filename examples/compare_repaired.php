<?php

declare(strict_types=1);

require_once __DIR__ . '/../vendor/autoload.php';

echo "🔍 COMPARAISON : Original vs Réparé par PowerPoint\n";
echo str_repeat('=', 70) . "\n\n";

$originalFile = $argv[1] ?? __DIR__ . '/../test_merge_powerpoint.pptx';
$repairedFile = $argv[2] ?? __DIR__ . '/../test_merge_powerpoint_repaired.pptx';

if (!file_exists($originalFile)) {
    echo "❌ Fichier original non trouvé: $originalFile\n";
    exit(1);
}

if (!file_exists($repairedFile)) {
    echo "❌ Fichier réparé non trouvé: $repairedFile\n";
    echo "💡 Veuillez réparer le fichier avec PowerPoint et le sauvegarder sous:\n";
    echo "   $repairedFile\n";
    exit(1);
}

$original = new ZipArchive();
$original->open($originalFile);

$repaired = new ZipArchive();
$repaired->open($repairedFile);

// Compare file lists
echo "📁 Comparaison de la structure des fichiers:\n";
$originalFiles = [];
for ($i = 0; $i < $original->numFiles; $i++) {
    $originalFiles[] = $original->getNameIndex($i);
}

$repairedFiles = [];
for ($i = 0; $i < $repaired->numFiles; $i++) {
    $repairedFiles[] = $repaired->getNameIndex($i);
}

sort($originalFiles);
sort($repairedFiles);

$removed = array_diff($originalFiles, $repairedFiles);
$added = array_diff($repairedFiles, $originalFiles);

if (empty($removed) && empty($added)) {
    echo "   ✓ Même nombre de fichiers (" . count($originalFiles) . ")\n";
} else {
    if (!empty($removed)) {
        echo "   ❌ Fichiers supprimés par PowerPoint:\n";
        foreach ($removed as $file) {
            echo "      - $file\n";
        }
    }
    if (!empty($added)) {
        echo "   ✅ Fichiers ajoutés par PowerPoint:\n";
        foreach ($added as $file) {
            echo "      + $file\n";
        }
    }
}
echo "\n";

// Compare XML files content
echo "📋 Comparaison du contenu des fichiers XML:\n";
$commonFiles = array_intersect($originalFiles, $repairedFiles);
$xmlDifferences = [];

foreach ($commonFiles as $file) {
    if (!str_ends_with($file, '.xml') && !str_ends_with($file, '.rels')) {
        continue;
    }
    
    $originalContent = $original->getFromName($file);
    $repairedContent = $repaired->getFromName($file);
    
    if ($originalContent === $repairedContent) {
        continue;
    }
    
    // Files are different
    $xmlDifferences[$file] = [
        'original' => $originalContent,
        'repaired' => $repairedContent,
    ];
}

if (empty($xmlDifferences)) {
    echo "   ✓ Aucune différence dans les fichiers XML\n";
} else {
    echo "   ⚠️  " . count($xmlDifferences) . " fichiers XML modifiés:\n";
    foreach (array_keys($xmlDifferences) as $file) {
        echo "      - $file\n";
    }
}
echo "\n";

// Detailed analysis of each modified file
if (!empty($xmlDifferences)) {
    echo "🔬 ANALYSE DÉTAILLÉE DES MODIFICATIONS:\n";
    echo str_repeat('=', 70) . "\n\n";
    
    foreach ($xmlDifferences as $file => $contents) {
        echo "📄 $file\n";
        echo str_repeat('-', 70) . "\n";
        
        // Pretty print XML for comparison
        $originalXml = simplexml_load_string($contents['original']);
        $repairedXml = simplexml_load_string($contents['repaired']);
        
        if ($originalXml === false || $repairedXml === false) {
            echo "   ❌ Erreur de parsing XML\n\n";
            continue;
        }
        
        // Format XML for comparison
        $dom1 = new DOMDocument('1.0');
        $dom1->preserveWhiteSpace = false;
        $dom1->formatOutput = true;
        $dom1->loadXML($contents['original']);
        $originalFormatted = $dom1->saveXML();
        
        $dom2 = new DOMDocument('1.0');
        $dom2->preserveWhiteSpace = false;
        $dom2->formatOutput = true;
        $dom2->loadXML($contents['repaired']);
        $repairedFormatted = $dom2->saveXML();
        
        // Simple line-by-line comparison
        $originalLines = explode("\n", $originalFormatted);
        $repairedLines = explode("\n", $repairedFormatted);
        
        $maxLines = max(count($originalLines), count($repairedLines));
        $diffsFound = 0;
        
        for ($i = 0; $i < $maxLines && $diffsFound < 10; $i++) {
            $origLine = $originalLines[$i] ?? '';
            $repLine = $repairedLines[$i] ?? '';
            
            if (trim($origLine) !== trim($repLine)) {
                $diffsFound++;
                echo "\n   Ligne " . ($i + 1) . ":\n";
                echo "   AVANT: " . trim($origLine) . "\n";
                echo "   APRÈS: " . trim($repLine) . "\n";
            }
        }
        
        if ($diffsFound >= 10) {
            echo "\n   ... (plus de différences non affichées)\n";
        }
        
        if ($diffsFound === 0) {
            echo "   ℹ️  Différences mineures (whitespace, formatage)\n";
        }
        
        echo "\n";
    }
}

// Specific checks
echo "🔍 VÉRIFICATIONS SPÉCIFIQUES:\n";
echo str_repeat('=', 70) . "\n\n";

// Check presentation.xml
echo "📊 presentation.xml:\n";
$origPres = $original->getFromName('ppt/presentation.xml');
$repPres = $repaired->getFromName('ppt/presentation.xml');

if ($origPres && $repPres) {
    $origPresXml = simplexml_load_string($origPres);
    $repPresXml = simplexml_load_string($repPres);
    
    $origPresXml->registerXPathNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
    $repPresXml->registerXPathNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
    
    // Compare slide IDs
    $origSlides = $origPresXml->xpath('//p:sldId');
    $repSlides = $repPresXml->xpath('//p:sldId');
    
    echo "   Slides (original): ";
    foreach ($origSlides as $slide) {
        echo (string)$slide['id'] . " ";
    }
    echo "\n";
    
    echo "   Slides (réparé):   ";
    foreach ($repSlides as $slide) {
        echo (string)$slide['id'] . " ";
    }
    echo "\n";
}
echo "\n";

// Check app.xml
echo "📋 docProps/app.xml:\n";
$origApp = $original->getFromName('docProps/app.xml');
$repApp = $repaired->getFromName('docProps/app.xml');

if ($origApp && $repApp) {
    $origAppXml = simplexml_load_string($origApp);
    $repAppXml = simplexml_load_string($repApp);
    
    echo "   Original - Slides: " . ((string)$origAppXml->Slides ?: 'N/A') . ", Notes: " . ((string)$origAppXml->Notes ?: 'N/A') . "\n";
    echo "   Réparé   - Slides: " . ((string)$repAppXml->Slides ?: 'N/A') . ", Notes: " . ((string)$repAppXml->Notes ?: 'N/A') . "\n";
}
echo "\n";

$original->close();
$repaired->close();

echo str_repeat('=', 70) . "\n";
echo "✅ Comparaison terminée\n";
echo "\n💡 Analysez les différences ci-dessus pour identifier ce que PowerPoint a corrigé.\n";