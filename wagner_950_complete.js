const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
        BorderStyle, WidthType, ShadingType, HeadingLevel, PageBreak, PageOrientation } = require('docx');
const PDFDocument = require('pdfkit');

// Species data organized alphabetically with their specific gravity settings and scientific names
const speciesData = [
  { name: "African Blackwood", setting: "1.27", scientific: "Dalbergia melanoxylon" },
  { name: "Afrormosia", setting: "0.65", scientific: "Pericopsis elata" },
  { name: "Alder, Red", setting: "0.41", scientific: "Alnus rubra" },
  { name: "American Red Oak", setting: "0.63", scientific: "Quercus rubra" },
  { name: "Andiroba", setting: "0.57", scientific: "Carapa guianensis" },
  { name: "Ash, Black", setting: "0.49", scientific: "Fraxinus nigra" },
  { name: "Ash, Blue", setting: "0.58", scientific: "Fraxinus quadrangulata" },
  { name: "Ash, Green", setting: "0.56", scientific: "Fraxinus pennsylvanica" },
  { name: "Ash, Oregon", setting: "0.55", scientific: "Fraxinus latifolia" },
  { name: "Ash, Red", setting: "0.55", scientific: "Fraxinus pennsylvanica" },
  { name: "Ash, White", setting: "0.60", scientific: "Fraxinus americana" },
  { name: "Aspen, Bigtooth", setting: "0.39", scientific: "Populus grandidentata" },
  { name: "Aspen, Quaking", setting: "0.38", scientific: "Populus tremuloides" },
  { name: "Avodire", setting: "0.51", scientific: "Turraeanthus africanus" },
  { name: "Baldcypress", setting: "0.46", scientific: "Taxodium distichum" },
  { name: "Balsa", setting: "0.14", scientific: "Ochroma pyramidale" },
  { name: "Balsamo (Myroxylon)", setting: "0.83", scientific: "Myroxylon balsamum" },
  { name: "Balsamo (Protium)", setting: "0.55", scientific: "Protium spp." },
  { name: "Banak (Virola)", setting: "0.45", scientific: "Virola spp." },
  { name: "Basswood, American", setting: "0.37", scientific: "Tilia americana" },
  { name: "Beech, American", setting: "0.64", scientific: "Fagus grandifolia" },
  { name: "Beech, Euro", setting: "0.67", scientific: "Fagus sylvatica" },
  { name: "Benge", setting: "0.70", scientific: "Guibourtia arnoldiana" },
  { name: "Birch, Paper", setting: "0.55", scientific: "Betula papyrifera" },
  { name: "Birch, Sweet", setting: "0.65", scientific: "Betula lenta" },
  { name: "Birch, White", setting: "0.53", scientific: "Betula populifolia" },
  { name: "Birch, Yellow", setting: "0.62", scientific: "Betula alleghaniensis" },
  { name: "Box", setting: "0.83", scientific: "Buxus sempervirens" },
  { name: "Brazilian Cherry", setting: "0.83", scientific: "Hymenaea courbaril" },
  { name: "Brazilian Mahogany", setting: "0.47", scientific: "Swietenia macrophylla" },
  { name: "British Elm", setting: "0.53", scientific: "Ulmus procera" },
  { name: "Bubinga", setting: "0.75", scientific: "Guibourtia spp." },
  { name: "Butternut", setting: "0.38", scientific: "Juglans cinerea" },
  { name: "Cativo", setting: "0.42", scientific: "Prioria copaifera" },
  { name: "Cedar, Alaska", setting: "0.44", scientific: "Callitropsis nootkatensis" },
  { name: "Cedar, Atlantic White", setting: "0.32", scientific: "Chamaecyparis thyoides" },
  { name: "Cedar, Eastern Red", setting: "0.47", scientific: "Juniperus virginiana" },
  { name: "Cedar, Incense", setting: "0.37", scientific: "Calocedrus decurrens" },
  { name: "Cedar, Northern White", setting: "0.31", scientific: "Thuja occidentalis" },
  { name: "Cedar of Lebanon", setting: "0.53", scientific: "Cedrus libani" },
  { name: "Cedar, Port Orford", setting: "0.43", scientific: "Chamaecyparis lawsoniana" },
  { name: "Cedar, Western Red", setting: "0.32", scientific: "Thuja plicata" },
  { name: "Cedar, Yellow", setting: "0.44", scientific: "Callitropsis nootkatensis" },
  { name: "Cedrella", setting: "0.39", scientific: "Cedrela odorata" },
  { name: "Cherry, Black", setting: "0.50", scientific: "Prunus serotina" },
  { name: "Chestnut, American", setting: "0.43", scientific: "Castanea dentata" },
  { name: "Cocobolo", setting: "0.85", scientific: "Dalbergia retusa" },
  { name: "Cottonwood, Balsam Poplar", setting: "0.34", scientific: "Populus balsamifera" },
  { name: "Cottonwood, Black", setting: "0.35", scientific: "Populus trichocarpa" },
  { name: "Cottonwood, Eastern", setting: "0.40", scientific: "Populus deltoides" },
  { name: "Degame", setting: "0.72", scientific: "Calycophyllum candidissimum" },
  { name: "Determa", setting: "0.55", scientific: "Ocotea rubra" },
  { name: "Dogwood, Flowering", setting: "0.72", scientific: "Cornus florida" },
  { name: "Douglas Fir", setting: "0.48", scientific: "Pseudotsuga menziesii" },
  { name: "Ebony", setting: "0.94", scientific: "Diospyros spp." },
  { name: "Elliotis Pine", setting: "0.59", scientific: "Pinus elliottii var. elliottii" },
  { name: "Elm, American", setting: "0.50", scientific: "Ulmus americana" },
  { name: "Elm, Rock", setting: "0.63", scientific: "Ulmus thomasii" },
  { name: "Elm, Slippery", setting: "0.53", scientific: "Ulmus rubra" },
  { name: "English Cherry", setting: "0.58", scientific: "Prunus avium" },
  { name: "English Oak", setting: "0.57", scientific: "Quercus robur" },
  { name: "European Ash", setting: "0.58", scientific: "Fraxinus excelsior" },
  { name: "European Hornbeam", setting: "0.74", scientific: "Carpinus betulus" },
  { name: "European Walnut", setting: "0.56", scientific: "Juglans regia" },
  { name: "Fir, Balsam", setting: "0.35", scientific: "Abies balsamea" },
  { name: "Fir, California Red", setting: "0.38", scientific: "Abies magnifica" },
  { name: "Fir, Grand", setting: "0.37", scientific: "Abies grandis" },
  { name: "Fir, Noble", setting: "0.39", scientific: "Abies procera" },
  { name: "Fir, Pacific Silver", setting: "0.43", scientific: "Abies amabilis" },
  { name: "Fir, Subalpine", setting: "0.32", scientific: "Abies lasiocarpa" },
  { name: "Fir, White", setting: "0.39", scientific: "Abies concolor" },
  { name: "Gombeira", setting: "1.00", scientific: "Didelotia africana" },
  { name: "Guatambu (Argentinean)", setting: "0.70", scientific: "Balfourodendron riedelianum" },
  { name: "Guatambu (Brazil)", setting: "0.79", scientific: "Aspidosperma spp." },
  { name: "Gum, Black", setting: "0.50", scientific: "Nyssa sylvatica" },
  { name: "Gum, Red", setting: "0.52", scientific: "Liquidambar styraciflua" },
  { name: "Hackberry", setting: "0.53", scientific: "Celtis occidentalis" },
  { name: "Hemlock, Eastern", setting: "0.40", scientific: "Tsuga canadensis" },
  { name: "Hemlock, Mountain", setting: "0.45", scientific: "Tsuga mertensiana" },
  { name: "Hemlock, Western", setting: "0.45", scientific: "Tsuga heterophylla" },
  { name: "Hickory (Pecan), Bitternut", setting: "0.66", scientific: "Carya cordiformis" },
  { name: "Hickory (Pecan), Nutmeg", setting: "0.60", scientific: "Carya myristiciformis" },
  { name: "Hickory (Pecan), Water", setting: "0.62", scientific: "Carya aquatica" },
  { name: "Hickory (True), Mockernut", setting: "0.72", scientific: "Carya tomentosa" },
  { name: "Hickory (True), Pignut", setting: "0.75", scientific: "Carya glabra" },
  { name: "Hickory (True), Shagbark", setting: "0.72", scientific: "Carya ovata" },
  { name: "Hickory (True), Shellbark", setting: "0.69", scientific: "Carya laciniosa" },
  { name: "Hickory, Pecan", setting: "0.66", scientific: "Carya illinoinensis" },
  { name: "Holly, American", setting: "0.55", scientific: "Ilex opaca" },
  { name: "Hophornbeam, Eastern", setting: "0.70", scientific: "Ostrya virginiana" },
  { name: "Hura", setting: "0.40", scientific: "Hura crepitans" },
  { name: "Indian Laurel", setting: "0.79", scientific: "Terminalia tomentosa" },
  { name: "Ipe", setting: "0.99", scientific: "Handroanthus spp." },
  { name: "Iroko", setting: "0.57", scientific: "Milicia excelsa" },
  { name: "Jacaranda", setting: "0.34", scientific: "Jacaranda mimosifolia" },
  { name: "Jarrah", setting: "0.75", scientific: "Eucalyptus marginata" },
  { name: "Jelutong", setting: "0.38", scientific: "Dyera costulata" },
  { name: "Kapur", setting: "0.70", scientific: "Dryobalanops spp." },
  { name: "Karri", setting: "0.79", scientific: "Eucalyptus diversicolor" },
  { name: "Keruing", setting: "0.76", scientific: "Dipterocarpus spp." },
  { name: "Kingwood", setting: "1.16", scientific: "Dalbergia cearensis" },
  { name: "KOA (Acacia Koa)", setting: "0.63", scientific: "Acacia koa" },
  { name: "Larch, Euro", setting: "0.48", scientific: "Larix decidua" },
  { name: "Larch, Western", setting: "0.52", scientific: "Larix occidentalis" },
  { name: "Laurel, California", setting: "0.55", scientific: "Umbellularia californica" },
  { name: "Lignum Vitae", setting: "1.13", scientific: "Guaiacum officinale" },
  { name: "Limba", setting: "0.40", scientific: "Terminalia superba" },
  { name: "Locust, Black", setting: "0.69", scientific: "Robinia pseudoacacia" },
  { name: "Macassar Ebony", setting: "0.90", scientific: "Diospyros celebica" },
  { name: "Madrone, Pacific", setting: "0.64", scientific: "Arbutus menziesii" },
  { name: "Magnolia, Southern", setting: "0.50", scientific: "Magnolia grandiflora" },
  { name: "Mahogany, African", setting: "0.44", scientific: "Khaya spp." },
  { name: "Mahogany, True", setting: "0.47", scientific: "Swietenia macrophylla" },
  { name: "Manni", setting: "0.63", scientific: "Symphonia globulifera" },
  { name: "Maple, Bigleaf", setting: "0.48", scientific: "Acer macrophyllum" },
  { name: "Maple, Black", setting: "0.57", scientific: "Acer nigrum" },
  { name: "Maple, Hard", setting: "0.60", scientific: "Acer saccharum" },
  { name: "Maple, Red", setting: "0.54", scientific: "Acer rubrum" },
  { name: "Maple, Silver", setting: "0.47", scientific: "Acer saccharinum" },
  { name: "Maple, Soft", setting: "0.49", scientific: "Acer rubrum" },
  { name: "Maple, Sugar", setting: "0.63", scientific: "Acer saccharum" },
  { name: "Merbau", setting: "0.67", scientific: "Intsia spp." },
  { name: "Mersawa", setting: "0.54", scientific: "Anisoptera spp." },
  { name: "Mesquite", setting: "0.86", scientific: "Prosopis spp." },
  { name: "Monkeypod", setting: "0.50", scientific: "Samanea saman" },
  { name: "Mountain Ash (Eucalyptus)", setting: "0.62", scientific: "Eucalyptus regnans" },
  { name: "Muninga", setting: "0.59", scientific: "Pterocarpus angolensis" },
  { name: "Myrtle, Oregon", setting: "0.55", scientific: "Umbellularia californica" },
  { name: "Myrtle, Tasmanian", setting: "0.64", scientific: "Nothofagus cunninghamii" },
  { name: "Oak (Red), Black", setting: "0.61", scientific: "Quercus velutina" },
  { name: "Oak (Red), Cherrybark", setting: "0.68", scientific: "Quercus pagoda" },
  { name: "Oak (Red), Laurel", setting: "0.63", scientific: "Quercus laurifolia" },
  { name: "Oak (Red), Northern", setting: "0.63", scientific: "Quercus rubra" },
  { name: "Oak (Red), Pin", setting: "0.63", scientific: "Quercus palustris" },
  { name: "Oak (Red), Scarlet", setting: "0.67", scientific: "Quercus coccinea" },
  { name: "Oak (Red), Southern", setting: "0.59", scientific: "Quercus falcata" },
  { name: "Oak (Red), Water", setting: "0.63", scientific: "Quercus nigra" },
  { name: "Oak (Red), Willow", setting: "0.69", scientific: "Quercus phellos" },
  { name: "Oak (White), Bur", setting: "0.64", scientific: "Quercus macrocarpa" },
  { name: "Oak (White), Chestnut", setting: "0.66", scientific: "Quercus prinus" },
  { name: "Oak (White), Overcup", setting: "0.63", scientific: "Quercus lyrata" },
  { name: "Oak (White), Post", setting: "0.67", scientific: "Quercus stellata" },
  { name: "Oak (White), Swamp", setting: "0.72", scientific: "Quercus bicolor" },
  { name: "Oak (White), Swamp Chestnut", setting: "0.67", scientific: "Quercus michauxii" },
  { name: "Oak, California Black", setting: "0.53", scientific: "Quercus kelloggii" },
  { name: "Oak, White", setting: "0.68", scientific: "Quercus alba" },
  { name: "Obeche", setting: "0.32", scientific: "Triplochiton scleroxylon" },
  { name: "Okoume", setting: "0.35", scientific: "Aucoumea klaineana" },
  { name: "Olive", setting: "0.81", scientific: "Olea europaea" },
  { name: "Opepe", setting: "0.68", scientific: "Nauclea diderrichii" },
  { name: "Padauk (P. indicus)", setting: "0.57", scientific: "Pterocarpus indicus" },
  { name: "Padauk (P. macrocarpus)", setting: "0.79", scientific: "Pterocarpus macrocarpus" },
  { name: "Padauk (P. marsupium)", setting: "0.71", scientific: "Pterocarpus marsupium" },
  { name: "Parana Pine", setting: "0.49", scientific: "Araucaria angustifolia" },
  { name: "Pecan", setting: "0.60", scientific: "Carya illinoinensis" },
  { name: "Peroba de Campos", setting: "0.66", scientific: "Paratecoma peroba" },
  { name: "Peroba Rosa", setting: "0.71", scientific: "Aspidosperma peroba" },
  { name: "Persimmon, Common", setting: "0.71", scientific: "Diospyros virginiana" },
  { name: "Pine, Eastern White", setting: "0.35", scientific: "Pinus strobus" },
  { name: "Pine, Hoop", setting: "0.44", scientific: "Araucaria cunninghamii" },
  { name: "Pine, Jack", setting: "0.43", scientific: "Pinus banksiana" },
  { name: "Pine, Loblolly", setting: "0.51", scientific: "Pinus taeda" },
  { name: "Pine, Lodgepole", setting: "0.41", scientific: "Pinus contorta" },
  { name: "Pine, Longleaf", setting: "0.59", scientific: "Pinus palustris" },
  { name: "Pine, Pitch", setting: "0.52", scientific: "Pinus rigida" },
  { name: "Pine, Pond", setting: "0.56", scientific: "Pinus serotina" },
  { name: "Pine, Ponderosa", setting: "0.40", scientific: "Pinus ponderosa" },
  { name: "Pine, Red", setting: "0.46", scientific: "Pinus resinosa" },
  { name: "Pine, Sand", setting: "0.48", scientific: "Pinus clausa" },
  { name: "Pine, Scots", setting: "0.45", scientific: "Pinus sylvestris" },
  { name: "Pine, Shortleaf", setting: "0.51", scientific: "Pinus echinata" },
  { name: "Pine, Slash", setting: "0.59", scientific: "Pinus elliottii" },
  { name: "Pine, Spruce", setting: "0.44", scientific: "Pinus glabra" },
  { name: "Pine, Sugar", setting: "0.36", scientific: "Pinus lambertiana" },
  { name: "Pine, Virginia", setting: "0.48", scientific: "Pinus virginiana" },
  { name: "Pine, Western White", setting: "0.35", scientific: "Pinus monticola" },
  { name: "Plane (Lacewood)", setting: "0.49", scientific: "Platanus spp." },
  { name: "Poplar, Yellow", setting: "0.42", scientific: "Liriodendron tulipifera" },
  { name: "Primavera", setting: "0.42", scientific: "Cybistax donnell-smithii" },
  { name: "Purpleheart", setting: "0.71", scientific: "Peltogyne spp." },
  { name: "Radiata Pine", setting: "0.45", scientific: "Pinus radiata" },
  { name: "Ramin", setting: "0.56", scientific: "Gonystylus spp." },
  { name: "Redwood, Old-Growth", setting: "0.40", scientific: "Sequoia sempervirens" },
  { name: "Redwood, Young-Growth", setting: "0.35", scientific: "Sequoia sempervirens" },
  { name: "Roble (Tabebuia)", setting: "0.55", scientific: "Tabebuia spp." },
  { name: "Rosewood, Brazilian", setting: "0.84", scientific: "Dalbergia nigra" },
  { name: "Rosewood, Indian", setting: "0.79", scientific: "Dalbergia latifolia" },
  { name: "Rubberwood", setting: "0.51", scientific: "Hevea brasiliensis" },
  { name: "Sapele", setting: "0.60", scientific: "Entandrophragma cylindricum" },
  { name: "Sassafras", setting: "0.46", scientific: "Sassafras albidum" },
  { name: "Spanish Cedar", setting: "0.44", scientific: "Cedrela odorata" },
  { name: "Spruce, Black", setting: "0.42", scientific: "Picea mariana" },
  { name: "Spruce, Engelmann", setting: "0.35", scientific: "Picea engelmannii" },
  { name: "Spruce, Northern", setting: "0.36", scientific: "Picea glauca" },
  { name: "Spruce, Red", setting: "0.40", scientific: "Picea rubens" },
  { name: "Spruce, Sitka", setting: "0.40", scientific: "Picea sitchensis" },
  { name: "Spruce, White", setting: "0.36", scientific: "Picea glauca" },
  { name: "Sweet Chestnut", setting: "0.51", scientific: "Castanea sativa" },
  { name: "Sweetgum", setting: "0.52", scientific: "Liquidambar styraciflua" },
  { name: "Sycamore, American", setting: "0.49", scientific: "Platanus occidentalis" },
  { name: "SYP (Southern Yellow Pine)", setting: "0.56", scientific: "Pinus spp." },
  { name: "Tamarack", setting: "0.53", scientific: "Larix laricina" },
  { name: "Tanoak", setting: "0.64", scientific: "Notholithocarpus densiflorus" },
  { name: "Tatajuba", setting: "0.72", scientific: "Bagassa guianensis" },
  { name: "Tauari (Couratari)", setting: "0.53", scientific: "Couratari spp." },
  { name: "Tawa (Beilschmiedia)", setting: "0.62", scientific: "Beilschmiedia tawa" },
  { name: "Tawa (Pometia)", setting: "0.58", scientific: "Pometia spp." },
  { name: "Teak", setting: "0.57", scientific: "Tectona grandis" },
  { name: "Tupelo, Black", setting: "0.50", scientific: "Nyssa sylvatica" },
  { name: "Tupelo, Water", setting: "0.50", scientific: "Nyssa aquatica" },
  { name: "Virola", setting: "0.45", scientific: "Virola spp." },
  { name: "Walnut, Black", setting: "0.55", scientific: "Juglans nigra" },
  { name: "Wenge", setting: "0.82", scientific: "Millettia laurentii" },
  { name: "Willow, Black", setting: "0.39", scientific: "Salix nigra" },
  { name: "Yellow-Poplar", setting: "0.42", scientific: "Liriodendron tulipifera" },
  { name: "Yew", setting: "0.63", scientific: "Taxus baccata" },
  { name: "Zebrano", setting: "0.77", scientific: "Microberlinia brazzavillensis" },
  { name: "Ziricote", setting: "0.81", scientific: "Cordia dodecandra" }
];

// Engineered materials
const engineeredMaterials = [
  { name: "Plywood", setting: "0.57", scientific: "Various species" },
  { name: "OSB", setting: "0.62", scientific: "Various species" },
  { name: "Permacore MDF", setting: "0.70", scientific: "Various species" },
  { name: "HDF Core", setting: "0.85", scientific: "Various species" },
  { name: "Advantech™", setting: "0.70", scientific: "Various species" }
];

// Common border styling
const tableBorder = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const cellBorders = { 
  top: tableBorder, 
  bottom: tableBorder, 
  left: tableBorder, 
  right: tableBorder 
};

// Create header cell
function createHeaderCell(text) {
  return new TableCell({
    borders: cellBorders,
    shading: { fill: "2E5C8A", type: ShadingType.CLEAR },
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: text, bold: true, color: "FFFFFF", size: 18 })]
    })]
  });
}

// Create data cell
function createDataCell(text, bold = false) {
  return new TableCell({
    borders: cellBorders,
    children: [new Paragraph({
      children: [new TextRun({ text: text, bold: bold, size: 16 })]
    })]
  });
}

// Split species into pages (30 species per page for readability)
function chunkArray(array, size) {
  const chunks = [];
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }
  return chunks;
}

// Create a table for species data with 3 columns
function createSpeciesTable(speciesChunk) {
  const rows = [
    new TableRow({
      tableHeader: true,
      children: [
        createHeaderCell("Common Name"),
        createHeaderCell("Scientific Name"),
        createHeaderCell("Setting")
      ]
    })
  ];

  speciesChunk.forEach(species => {
    rows.push(new TableRow({
      children: [
        createDataCell(species.name),
        createDataCell(species.scientific),
        createDataCell(species.setting, true)
      ]
    }));
  });

  return new Table({
    columnWidths: [3120, 4680, 1560],
    width: { size: 100, type: WidthType.PERCENTAGE },
    margins: { top: 50, bottom: 50, left: 100, right: 100 },
    rows: rows
  });
}

// Build document sections
const sections = [];

// Title page
sections.push({
  properties: {
    page: {
      margin: { top: 720, right: 720, bottom: 720, left: 720 },
      size: { orientation: PageOrientation.PORTRAIT }
    }
  },
  children: [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 2880, after: 480 },
      children: [new TextRun({ 
        text: "Wagner Orion 950", 
        bold: true, 
        size: 56,
        font: "Arial"
      })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 480 },
      children: [new TextRun({ 
        text: "Moisture Meter", 
        bold: true, 
        size: 48,
        font: "Arial"
      })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 1440 },
      children: [new TextRun({ 
        text: "Species Settings Reference", 
        size: 36,
        font: "Arial"
      })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 1440, after: 240 },
      border: {
        top: { style: BorderStyle.SINGLE, size: 12, color: "2E5C8A" },
        bottom: { style: BorderStyle.SINGLE, size: 12, color: "2E5C8A" }
      },
      children: [new TextRun({ 
        text: "Quick Reference Guide", 
        bold: true, 
        size: 28,
        color: "2E5C8A",
        font: "Arial"
      })]
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 960, after: 120 },
      children: [new TextRun({ 
        text: "How to Use:", 
        bold: true, 
        size: 24,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "1. Locate your wood species in the alphabetical tables", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "2. Note the specific gravity (SG) setting value", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "3. Enter this value into your Wagner 950 meter using the up/down arrows", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { before: 480, after: 120 },
      children: [new TextRun({ 
        text: "Important Notes:", 
        bold: true, 
        size: 24,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Settings are based on specific gravity at 12% moisture content", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Values represent average density for each species", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Natural variation of ±10% is normal within species", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Scientific names help ensure precise wood identification", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { before: 960 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ 
        text: "Wagner Meters • www.wagnermeters.com • (800) 634-9961", 
        size: 18,
        color: "666666",
        font: "Arial"
      })]
    })
  ]
});

// Split species data into chunks of 30 per page
const speciesChunks = chunkArray(speciesData, 30);

speciesChunks.forEach((chunk, index) => {
  const pageChildren = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 360 },
      children: [new TextRun({ 
        text: `Wood Species Settings (Page ${index + 1} of ${speciesChunks.length})`, 
        bold: true, 
        size: 28,
        color: "2E5C8A",
        font: "Arial"
      })]
    }),
    createSpeciesTable(chunk)
  ];

  sections.push({
    properties: {
      page: {
        margin: { top: 720, right: 720, bottom: 720, left: 720 }
      }
    },
    children: pageChildren
  });
});

// Engineered materials page
sections.push({
  properties: {
    page: {
      margin: { top: 720, right: 720, bottom: 720, left: 720 }
    }
  },
  children: [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 360 },
      children: [new TextRun({ 
        text: "Engineered Wood Products", 
        bold: true, 
        size: 28,
        color: "2E5C8A",
        font: "Arial"
      })]
    }),
    createSpeciesTable(engineeredMaterials),
    new Paragraph({
      spacing: { before: 720, after: 120 },
      children: [new TextRun({ 
        text: "Additional Information:", 
        bold: true, 
        size: 22,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• For species not listed, visit Wagner's online database at:", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 240 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ 
        text: "www.wagnermeters.com/specific-gravity", 
        size: 20,
        bold: true,
        color: "2E5C8A",
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• The online database contains over 7,500 wood species", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { before: 480, after: 120 },
      children: [new TextRun({ 
        text: "Calibration:", 
        bold: true, 
        size: 22,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Use the included On-Demand Calibrator regularly to maintain accuracy", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ 
        text: "• Calibrate before important measurements or if meter has been dropped", 
        size: 20,
        font: "Arial"
      })]
    }),
    new Paragraph({
      spacing: { before: 720 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ 
        text: "For technical support: support@wagnermeters.com", 
        size: 18,
        color: "666666",
        font: "Arial"
      })]
    })
  ]
});

// PDF Generation Functions
function createPDF(outputPath) {
  const doc = new PDFDocument({
    size: 'LETTER',
    margins: { top: 50, bottom: 50, left: 50, right: 50 }
  });

  const stream = fs.createWriteStream(outputPath);
  doc.pipe(stream);

  // Title Page
  doc.fontSize(28).fillColor('#000000').font('Helvetica-Bold')
     .text('Wagner Orion 950', { align: 'center' })
     .moveDown(0.5);

  doc.fontSize(24).text('Moisture Meter', { align: 'center' })
     .moveDown(0.5);

  doc.fontSize(18).font('Helvetica')
     .text('Species Settings Reference', { align: 'center' })
     .moveDown(2);

  // Quick Reference Guide section
  doc.fontSize(14).fillColor('#2E5C8A').font('Helvetica-Bold')
     .text('Quick Reference Guide', { align: 'center' })
     .moveDown(1);

  doc.fontSize(12).fillColor('#000000').font('Helvetica-Bold')
     .text('How to Use:', { align: 'left' })
     .moveDown(0.3);

  doc.fontSize(10).font('Helvetica')
     .text('1. Locate your wood species in the alphabetical tables')
     .text('2. Note the specific gravity (SG) setting value')
     .text('3. Enter this value into your Wagner 950 meter using the up/down arrows')
     .moveDown(1);

  doc.fontSize(12).font('Helvetica-Bold')
     .text('Important Notes:')
     .moveDown(0.3);

  doc.fontSize(10).font('Helvetica')
     .text('• Settings are based on specific gravity at 12% moisture content')
     .text('• Values represent average density for each species')
     .text('• Natural variation of ±10% is normal within species')
     .text('• Scientific names help ensure precise wood identification')
     .moveDown(2);

  doc.fontSize(9).fillColor('#666666')
     .text('Wagner Meters • www.wagnermeters.com • (800) 634-9961', { align: 'center' });

  // Species Tables
  const speciesChunks = chunkArray(speciesData, 30);

  speciesChunks.forEach((chunk, index) => {
    doc.addPage();

    doc.fontSize(14).fillColor('#2E5C8A').font('Helvetica-Bold')
       .text(`Wood Species Settings (Page ${index + 1} of ${speciesChunks.length})`, { align: 'center' })
       .moveDown(0.5);

    drawSpeciesTable(doc, chunk);
  });

  // Engineered Materials Page
  doc.addPage();

  doc.fontSize(14).fillColor('#2E5C8A').font('Helvetica-Bold')
     .text('Engineered Wood Products', { align: 'center' })
     .moveDown(0.5);

  drawSpeciesTable(doc, engineeredMaterials);

  doc.moveDown(2);

  doc.fontSize(11).fillColor('#000000').font('Helvetica-Bold')
     .text('Additional Information:')
     .moveDown(0.3);

  doc.fontSize(10).font('Helvetica')
     .text('• For species not listed, visit Wagner\'s online database at:')
     .moveDown(0.3);

  doc.fontSize(10).fillColor('#2E5C8A').font('Helvetica-Bold')
     .text('www.wagnermeters.com/specific-gravity', { align: 'center' })
     .moveDown(0.5);

  doc.fontSize(10).fillColor('#000000').font('Helvetica')
     .text('• The online database contains over 7,500 wood species')
     .moveDown(1);

  doc.fontSize(11).font('Helvetica-Bold')
     .text('Calibration:')
     .moveDown(0.3);

  doc.fontSize(10).font('Helvetica')
     .text('• Use the included On-Demand Calibrator regularly to maintain accuracy')
     .text('• Calibrate before important measurements or if meter has been dropped')
     .moveDown(2);

  doc.fontSize(9).fillColor('#666666')
     .text('For technical support: support@wagnermeters.com', { align: 'center' });

  doc.end();

  return new Promise((resolve, reject) => {
    stream.on('finish', resolve);
    stream.on('error', reject);
  });
}

function drawSpeciesTable(doc, speciesArray) {
  const tableTop = doc.y;
  const colWidths = [180, 270, 90]; // Common Name, Scientific Name, Setting
  const rowHeight = 20;
  const headerColor = '#2E5C8A';

  // Draw header row
  let y = tableTop;
  const headers = ['Common Name', 'Scientific Name', 'Setting'];

  // Header background
  doc.rect(50, y, colWidths[0] + colWidths[1] + colWidths[2], rowHeight)
     .fillAndStroke(headerColor, '#000000');

  // Header text
  doc.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(9);
  doc.text(headers[0], 55, y + 6, { width: colWidths[0] - 10 });
  doc.text(headers[1], 55 + colWidths[0], y + 6, { width: colWidths[1] - 10 });
  doc.text(headers[2], 55 + colWidths[0] + colWidths[1], y + 6, { width: colWidths[2] - 10 });

  y += rowHeight;

  // Draw data rows
  doc.fillColor('#000000').font('Helvetica').fontSize(8);

  speciesArray.forEach((species, idx) => {
    // Alternate row backgrounds for readability
    if (idx % 2 === 0) {
      doc.rect(50, y, colWidths[0] + colWidths[1] + colWidths[2], rowHeight)
         .fillAndStroke('#F5F5F5', '#000000');
    } else {
      doc.rect(50, y, colWidths[0] + colWidths[1] + colWidths[2], rowHeight)
         .stroke('#000000');
    }

    // Draw cell borders
    doc.rect(50, y, colWidths[0], rowHeight).stroke('#000000');
    doc.rect(50 + colWidths[0], y, colWidths[1], rowHeight).stroke('#000000');
    doc.rect(50 + colWidths[0] + colWidths[1], y, colWidths[2], rowHeight).stroke('#000000');

    // Draw text
    doc.fillColor('#000000').font('Helvetica').fontSize(8);
    doc.text(species.name, 55, y + 6, { width: colWidths[0] - 10, lineBreak: false });
    doc.text(species.scientific, 55 + colWidths[0], y + 6, { width: colWidths[1] - 10, lineBreak: false });
    doc.font('Helvetica-Bold');
    doc.text(species.setting, 55 + colWidths[0] + colWidths[1], y + 6, { width: colWidths[2] - 10, lineBreak: false });

    y += rowHeight;
  });

  doc.y = y + 10;
}

// Create DOCX document
const docxDoc = new Document({ sections });

// Get command line arguments
const args = process.argv.slice(2);
const format = args[0] || 'both'; // 'docx', 'pdf', or 'both'

// Save to file(s)
async function generateDocuments() {
  try {
    if (format === 'docx' || format === 'both') {
      const buffer = await Packer.toBuffer(docxDoc);
      fs.writeFileSync("Wagner_950_Reference_Sheets.docx", buffer);
      console.log("✓ DOCX file created: Wagner_950_Reference_Sheets.docx");
    }

    if (format === 'pdf' || format === 'both') {
      await createPDF("Wagner_950_Reference_Sheets.pdf");
      console.log("✓ PDF file created: Wagner_950_Reference_Sheets.pdf");
    }

    if (format !== 'docx' && format !== 'pdf' && format !== 'both') {
      console.log("Usage: node wagner_950_complete.js [docx|pdf|both]");
      console.log("Default: both");
    } else {
      console.log("\nReference sheets created successfully with scientific names!");
    }
  } catch (error) {
    console.error("Error generating documents:", error);
  }
}

generateDocuments();
