/**
 * ================================================================================
 * GBP SERVICE DESCRIPTION GENERATOR
 * ================================================================================
 * Generates Google Business Profile service descriptions for Abbey Floor Care.
 * Output: Google Sheet with Service, Price, Service Description columns.
 * Each description is grounded in the TSM for that material.
 * Run function: generateGBPServiceSheet()
 * ================================================================================
 */

function generateGBPServiceSheet() {

  // ── SERVICE DATA ──────────────────────────────────────────────────────────────
  // Priority 1 = Named stone/tile, core service
  // Priority 2 = Named repair or treatment, material-adjacent
  // Priority 3 = Generic service
  // Remove = no relevance to specialist positioning

  var services = [

    // ── PRIORITY 1: Named stone or tile, core service ──────────────────────────

    {
      service: "Marble Floor Cleaning",
      priority: 1,
      price: "From £45/m²",
      description: "Marble floor cleaning removes soil, residues, and failed sealers using alkaline chemistry and controlled mechanical agitation. Marble's calcium carbonate structure reacts immediately with acidic products — all cleaning uses pH-neutral or alkaline chemistry only. Wet vacuum extraction prevents recontamination. Impregnating sealer applied on completion."
    },
    {
      service: "Marble Floor Polishing",
      priority: 1,
      price: "From £45/m²",
      description: "Marble floor polishing uses progressive diamond abrasive pads to remove micro-scratching and acid etching, rebuilding reflectivity through the calcite crystal structure. Honing removes the damaged surface layer; polishing refines it to a clear mirror finish. Spray-burnishing enhances final gloss clarity. Impregnating sealer applied on completion."
    },
    {
      service: "Marble Restoration",
      priority: 1,
      price: "From £55/m²",
      description: "Marble restoration corrects acid etching, micro-scratching, lippage, and surface wear using diamond abrasive sequences from coarse milling pads through to fine polishing. Cleaning removes residues and failed sealers first. Chips and cracks are filled with colour-matched resin. Impregnating sealer applied on completion to slow future staining."
    },
    {
      service: "Marble Cleaners",
      priority: 1,
      price: "From £45/m²",
      description: "Professional marble cleaning uses high-alkaline or pH-neutral chemistry to remove organic soil, grease, and failed sealers without damaging the calcium carbonate surface. Acidic products are absolutely prohibited — even mild household acids cause immediate and permanent etching. Mechanical agitation with soft pads and full slurry extraction ensure contamination is removed from the stone."
    },
    {
      service: "Marble Floor Repair",
      priority: 1,
      price: "POA",
      description: "Marble floor repair addresses chips, cracks, and localised surface damage using colour-matched resin fillers blended to the stone's natural veining. Repairs are ground and honed flush with the surrounding surface. Deep structural cracks require investigation before cosmetic filling. Where lippage exceeds 2mm, grinding is carried out before honing to achieve a flat plane."
    },
    {
      service: "Marble Polishing",
      priority: 1,
      price: "From £45/m²",
      description: "Marble polishing uses progressive diamond pads and spray-burnishing to rebuild reflectivity after honing. The dense interlocking calcite crystal structure enables high mechanical polish — results vary by stone hardness and origin. Polished marble shows a clear mirror-like reflection under natural light. Impregnating sealer is applied on completion."
    },
    {
      service: "Marble Tiles",
      priority: 1,
      price: "From £45/m²",
      description: "Marble tile restoration addresses the full range of surface conditions including acid etching, micro-scratching, lippage, grout deterioration, and staining. Assessment establishes whether cleaning alone is sufficient or whether honing and polishing are required. Tile-by-tile condition mapping guides the intervention sequence. Impregnating sealer applied on completion."
    },
    {
      service: "Limestone Floor Cleaning",
      priority: 1,
      price: "From £40/m²",
      description: "Limestone floor cleaning uses alkaline chemistry to remove organic soil, old sealers, and embedded residues. Acidic cleaners are absolutely prohibited — limestone is calcium carbonate and etches immediately on acid contact. Where a slurry-weakened upper layer has developed from decades of mop-water exposure, honing may be recommended after cleaning to restore stain resistance."
    },
    {
      service: "Limestone Cleaning",
      priority: 1,
      price: "From £40/m²",
      description: "Professional limestone cleaning removes surface contamination, failed sealers, and the residue films that cause rapid re-soiling. Alkaline chemistry with mechanical agitation and full wet-vacuum extraction is used throughout. Limescale must be removed mechanically — acid descalers cause permanent surface dissolution. Impregnating sealer applied on completion."
    },
    {
      service: "Limestone Cleaners",
      priority: 1,
      price: "From £40/m²",
      description: "Limestone cleaning uses alkaline and pH-neutral chemistry to safely remove soil, grease, and coating residues without attacking the calcium carbonate surface. All acidic products are prohibited. Where older floors have developed a softened, absorbent surface layer from long-term mop-water exposure, honing is carried out to restore the substrate before sealing."
    },
    {
      service: "Limestone Cleaner",
      priority: 1,
      price: "From £40/m²",
      description: "Safe limestone cleaning requires chemistry that does not react with calcium carbonate. Alkaline degreasers remove organic soil and failed sealers; pH-neutral cleaners maintain the surface between professional visits. Mechanical agitation with soft brushes and immediate slurry extraction completes the process. Impregnating sealer applied to protect the restored surface."
    },
    {
      service: "Travertine Cleaning",
      priority: 1,
      price: "From £40/m²",
      description: "Travertine cleaning removes soil, residues, and failed sealers from within the stone's natural void network using controlled alkaline chemistry and full wet-vacuum extraction. Acidic cleaners are prohibited — travertine shares limestone's calcium carbonate sensitivity. All alkaline cleaning products must be fully rinsed to prevent residual alkalinity degrading the sealer system over time."
    },
    {
      service: "Travertine Floors Cleaning",
      priority: 1,
      price: "From £40/m²",
      description: "Professional travertine floor cleaning extracts soil and cleaning residues from natural voids and grout joints using alkaline chemistry, mechanical agitation, and pressurised extraction. The void structure traps contamination mechanically — surface mopping redistributes rather than removes it. Selective void filling and impregnating or topical sealer applied on completion."
    },
    {
      service: "Travertine Cleaned",
      priority: 1,
      price: "From £40/m²",
      description: "Travertine is cleaned using pH-neutral or alkaline chemistry to remove surface contamination and void-trapped soil. Full rinsing after any alkaline product is essential — residual alkalinity degrades sealers progressively. Wet-vacuum extraction removes dissolved contamination immediately. Impregnating or topical sealer applied on completion to protect the restored surface."
    },
    {
      service: "Travertine Cleaners",
      priority: 1,
      price: "From £40/m²",
      description: "Travertine cleaning uses alkaline and pH-neutral chemistry to safely remove organic soil, failed sealers, and void-trapped residues without damaging the calcium carbonate surface. Acidic products are absolutely prohibited. Professional hot water extraction removes contamination that domestic mopping redistributes into voids. Sealer applied on completion."
    },
    {
      service: "Slate Floor Cleaning",
      priority: 1,
      price: "From £35/m²",
      description: "Slate floor cleaning removes old acrylic sealers, wax build-up, ingrained soil, and cleaning residues using solvent-based strippers and alkaline degreasers matched to the slate's origin and finish. Riven slate requires rotary machines with appropriate brushes to reach recessed texture. Colour-enhancing impregnating sealer restores mineral pigment depth on completion."
    },
    {
      service: "Slate Cleaning",
      priority: 1,
      price: "From £35/m²",
      description: "Professional slate cleaning strips old coatings and removes ingrained soil from riven texture using solvent-based strippers and alkaline chemistry. Slate is sensitive to strongly acidic and strongly alkaline extremes — chemistry is matched to origin. Wet-vacuum extraction prevents re-soiling. Colour-enhancing impregnator or topical sealer applied on completion."
    },
    {
      service: "Slate Cleaner",
      priority: 1,
      price: "From £35/m²",
      description: "Slate cleaning uses pH-neutral or alkaline chemistry to safely remove surface residues and old coating build-up. Origin matters — Welsh, Indian, and Chinese slate all respond differently to chemistry and sealer systems. After thorough extraction and drying, colour-enhancing impregnating sealer is applied to restore mineral pigment depth and protect the surface."
    },
    {
      service: "Granite Cleaning",
      priority: 1,
      price: "From £35/m²",
      description: "Granite floor cleaning removes ingrained soil, failed sealers, and surface residues using alkaline chemistry and mechanical agitation. Granite is a silicate rock — it does not react to acidic cleaners in the way calcium carbonate stone does, but pH-neutral maintenance is still recommended to preserve sealers. Impregnating sealer applied on completion."
    },
    {
      service: "Granite Floors",
      priority: 1,
      price: "From £35/m²",
      description: "Granite floor restoration addresses ingrained soil, surface dulling, and failed sealers. Granite's silicate composition makes it more chemically resistant than calcium carbonate stone, but mechanical scratching dulls polished surfaces over time. Honing and polishing restores reflectivity where required. Impregnating sealer applied to protect the surface on completion."
    },
    {
      service: "Sandstone Sealer",
      priority: 1,
      price: "From £15/m²",
      description: "Sandstone sealing uses breathable impregnating sealers on all floors without a confirmed damp-proof membrane — non-breathable coatings trap rising moisture and cause whitening and peeling. Colour-enhancing impregnators deepen natural grain tone. The floor must be fully clean and dry before application. Sealer reduces soil absorption and improves day-to-day cleanability."
    },
    {
      service: "Terracotta Tile Cleaning",
      priority: 1,
      price: "From £35/m²",
      description: "Terracotta tile cleaning requires complete removal of all coatings, waxes, and oils before any new sealer can be applied — partial stripping leads to uneven appearance. High-alkaline strippers address acrylic coatings; solvent-based strippers remove polymerised linseed oil. The floor must be bone dry before sealing — typically 48–72 hours minimum. Breathable sealer applied on completion."
    },
    {
      service: "Terrazzo Floors",
      priority: 1,
      price: "From £50/m²",
      description: "Terrazzo floor restoration addresses binder degradation, embedded soil, and surface wear using mechanical grinding, honing, and polishing. Portland cement binder softens with age and alkaline cleaning, causing persistent dullness that cleaning cannot resolve. Grinding removes the degraded layer; honing refines it; sealing protects the renewed surface. Acidic cleaners are prohibited."
    },
    {
      service: "Natural Stone Floor Cleaning",
      priority: 1,
      price: "From £35/m²",
      description: "Natural stone floor cleaning uses chemistry and tooling matched to the specific material — the same treatment that restores one stone permanently damages another. Calcium carbonate stones (marble, limestone, travertine) require alkaline chemistry and prohibit all acids. Porous sedimentary stones require breathable sealing systems. Assessment before treatment determines the correct protocol."
    },
    {
      service: "Natural Stone Floor Restoration",
      priority: 1,
      price: "From £45/m²",
      description: "Natural stone floor restoration is a diagnostic-led process — material identification governs every decision. Marble and limestone require diamond honing to correct etching and surface wear. Sandstone and terracotta require breathable sealing systems without mechanical refinement. Slate requires origin-specific sealer selection. Assessment establishes the correct sequence before any work begins."
    },
    {
      service: "Natural Stone Tile Restoration",
      priority: 1,
      price: "From £45/m²",
      description: "Natural stone tile restoration addresses the full range of surface conditions including soil absorption, etching, micro-scratching, grout deterioration, and sealer failure. Treatment depends entirely on stone type — calcite-based stones require alkaline chemistry and diamond resurfacing; clay-based stones require breathable sealers and controlled extraction. Assessment determines the correct approach."
    },
    {
      service: "Natural Stone Cleaner",
      priority: 1,
      price: "From £35/m²",
      description: "Professional natural stone cleaning uses chemistry matched to the specific material. Calcium carbonate stones — marble, limestone, travertine, terrazzo — require alkaline or pH-neutral products only; acid contact causes immediate and irreversible etching. Siliceous and clay-based stones require breathable sealing on ground-bearing floors. Assessment identifies the stone and selects the correct cleaning system."
    },
    {
      service: "Natural Stone Floor Cleaner",
      priority: 1,
      price: "From £35/m²",
      description: "Safe natural stone cleaning requires material identification before chemistry selection. Alkaline degreasers remove organic soil from most stone types; pH-neutral cleaners maintain surfaces between professional visits. Acidic products are prohibited on calcium carbonate stones. Mechanical agitation with appropriate brushes and full slurry extraction ensure contamination is removed at depth."
    },
    {
      service: "Porcelain Tiles Cleaning",
      priority: 1,
      price: "From £25/m²",
      description: "Porcelain tile cleaning addresses surface film build-up, ingrained dirt in textured finishes, polymer grout haze, and soiled grout joints. The tile body is vitrified and non-porous — contamination accumulates on the surface or in grout. Grout recolouring delivers the most dramatic visual transformation. Polished porcelain can be burnished with ultra-fine diamond pads to restore gloss where dulled by micro-abrasion."
    },

    // ── PRIORITY 2: Named repair or treatment, material-adjacent ───────────────

    {
      service: "Stain Removal",
      priority: 2,
      price: "From £15/m²",
      description: "Stone and tile stain removal uses chemistry matched to both the stain type and the material. Organic stains on natural stone respond to alkaline gels or specialist reducing agents. Rust staining requires targeted acidic gel — prohibited on calcium carbonate stone where mechanical removal is used instead. Iron oxidation staining in marble and limestone may not fully lift. Assessment before treatment is essential."
    },
    {
      service: "Stain Removal Treating",
      priority: 2,
      price: "From £15/m²",
      description: "Professional stain treatment uses poultice extraction, alkaline gels, specialist reducing agents, or mechanical methods depending on stain type and stone composition. Acidic stain removers are prohibited on calcium carbonate stone — they cause further damage. Deep-set organic staining may require heat-activated extraction. Some stains will not fully lift; honest assessment before treatment manages expectations."
    },
    {
      service: "Grout Sealing",
      priority: 2,
      price: "From £10/m²",
      description: "Grout sealing applies an impregnating sealer, topical spray sealer, or epoxy grout colourant to cleaned grout joints to reduce future soil absorption. Epoxy grout colourant provides the highest protection — simultaneously restoring colour and creating a waterproof barrier. Grout must be fully cleaned and dry before any sealer application. All natural stone grout requires pH-neutral maintenance chemistry after sealing."
    },
    {
      service: "Grout Repair",
      priority: 2,
      price: "POA",
      description: "Grout repair addresses cracked, missing, or structurally failed joints. Failed grout is removed using oscillating tools with care taken to avoid edge damage to adjacent tiles. New grout is matched to the original colour and joint width. Flexible grout formulations are used on wooden subfloors. On historic clay tile floors, lime-based grout is used to maintain breathability."
    },
    {
      service: "Grout Repairs",
      priority: 2,
      price: "POA",
      description: "Professional grout repair removes cracked or failed grout and applies matched replacement. On natural stone floors, colour matching to the stone tone is important. On wooden subfloors, flexible formulations are mandatory to accommodate movement. Grout recolouring is preferred where grout is cosmetically deteriorated but physically sound — less disruptive and more durable than re-grouting."
    },
    {
      service: "Tile And Grout Cleaning",
      priority: 2,
      price: "From £25/m²",
      description: "Tile and grout cleaning uses dual-chemistry alkaline and acidic protocols to address the full range of contamination — organic soil in grout joints, polymer grout haze on tile surfaces, and mineral deposits. Polymer grout haze requires two-stage removal: alkaline first, acidic second. Full mechanical agitation and wet-vacuum extraction prevent re-soiling. Grout sealing applied on completion."
    },
    {
      service: "Tile Cleaners",
      priority: 2,
      price: "From £25/m²",
      description: "Professional tile cleaning matches chemistry to tile type — alkaline degreasers for organic soil, mild acid for grout haze and mineral deposits, solvent-based strippers for coating residues. On calcium carbonate stone tiles, acid is absolutely prohibited. Mechanical agitation penetrates grout joints and textured surfaces that routine mopping cannot reach. Full slurry extraction completes each cleaning pass."
    },
    {
      service: "Tiles Repair",
      priority: 2,
      price: "POA",
      description: "Tile repair covers chips, cracks, and missing tiles across stone and ceramic formats. Natural stone chips are filled with colour-matched resin blended to the stone's veining and tone. Ceramic and porcelain tiles cannot be repaired invisibly — replacement is the correct response. Victorian and Edwardian clay tiles are replaced with reclaimed period-matched originals to maintain heritage continuity."
    },
    {
      service: "Chip Repairs",
      priority: 2,
      price: "POA",
      description: "Stone chip repair uses colour-matched resin filler incorporating similar aggregate or pigment to blend with the surrounding surface. Repairs are ground and honed flush as part of the resurfacing sequence. Marble, limestone, travertine, and terrazzo can all be chip-repaired with reasonable results. Ceramic and porcelain chips require tile replacement — the glazed or vitrified body cannot be filled invisibly."
    },
    {
      service: "Cracks Repaired",
      priority: 2,
      price: "POA",
      description: "Stone floor crack repair uses colour-matched resin or cementitious filler depending on crack width and movement risk. Hairline cracks are generally left unless aesthetically required. Re-opening cracks signal subfloor movement requiring structural investigation before cosmetic repair. In terrazzo, service channel holes and historic fixing points are rebuilt level before grinding and honing."
    },
    {
      service: "Crack Repair",
      priority: 2,
      price: "POA",
      description: "Crack repair in stone and tile floors uses resin or lime-compatible filler depending on material type. Structural cracks indicating subfloor movement require investigation before filling. Cosmetic hairline cracks in natural stone are filled with colour-matched resin. Historic clay tile cracks are filled with lime mortar. All repairs are ground and honed flush with the surrounding surface."
    },
    {
      service: "Crack Repairs",
      priority: 2,
      price: "POA",
      description: "Professional crack repair addresses isolated and widespread fracture patterns in stone and tile floors. Isolated chips and surface cracks receive colour-matched resin fills. Widespread cracking patterns indicate substrate instability requiring structural assessment. In terrazzo, newly exposed pinholes during grinding are filled using cementitious grout float before honing proceeds."
    },
    {
      service: "Resin Repair",
      priority: 2,
      price: "POA",
      description: "Colour-matched resin repair addresses chips, cracks, and voids in natural stone floors including marble, limestone, travertine, slate, and terrazzo. Resin is mixed to match the stone's base colour and veining tone, applied, cured, and ground flush with the surrounding surface during the honing sequence. Deep voids in travertine are filled in staged pours to prevent sinkage."
    },
    {
      service: "Colour Enhancing",
      priority: 2,
      price: "From £10/m²",
      description: "Colour-enhancing impregnating sealer deepens natural mineral tone in stone and clay tile floors without forming a surface film. Particularly effective on limestone, sandstone, slate, terracotta, and Victorian clay tiles that appear washed-out after deep cleaning. Enhancement is permanent and must be agreed before application. Multiple coats are typically required on highly porous surfaces."
    },
    {
      service: "Epoxy Grout",
      priority: 2,
      price: "From £15/m²",
      description: "Epoxy grout colourant is applied over thoroughly cleaned existing grout to restore uniform colour, create a waterproof barrier, and dramatically reduce future porosity. It is the preferred protection method for all ceramic and porcelain tile restorations — simultaneously restoring appearance and providing maximum protection. Grout must be fully cleaned and pre-treated before colourant application."
    },
    {
      service: "Hole Filler",
      priority: 2,
      price: "POA",
      description: "Natural void and hole filling in travertine, terrazzo, and stone floors uses colour-matched resin or cementitious filler depending on void size and context. Travertine void filling is a core part of every restoration — targeted at visually problematic or structurally unstable areas. Terrazzo pinholes exposed during grinding are filled by cementitious grout float before honing. All fills are ground and honed flush."
    },
    {
      service: "Damage Repaired",
      priority: 2,
      price: "POA",
      description: "Stone and tile damage repair addresses chips, cracks, etching, and surface wear using the correct intervention for the material. Acid etching on marble and limestone requires mechanical honing — cleaning cannot correct chemical surface damage. Chips are filled with colour-matched resin. Grout damage is addressed by recolouring or re-grouting. Assessment before treatment establishes the correct repair approach."
    },
    {
      service: "Stone Sealer",
      priority: 2,
      price: "From £10/m²",
      description: "Stone sealing uses impregnating or topical systems matched to the stone type and moisture conditions. Calcium carbonate stones receive impregnating sealers that slow staining without preventing etching — acid contact still damages the surface below the sealer. Ground-bearing floors without a damp-proof membrane require breathable impregnating systems. Topical coatings are used where a specific sheen finish is required and moisture is confirmed stable."
    },
    {
      service: "Sealing Service",
      priority: 2,
      price: "From £10/m²",
      description: "Professional stone and tile sealing selects the correct system for the material, finish, and moisture conditions. Breathable impregnating sealers are mandatory on ground-bearing floors without a damp-proof membrane. Colour-enhancing impregnators deepen natural tone without surface film. Topical sealers add sheen on moisture-stable floors. The surface must be fully clean and dry before any sealer application."
    },
    {
      service: "Diamond Cutting",
      priority: 2,
      price: "POA",
      description: "Diamond abrasive tooling is used in natural stone restoration for progressive honing and polishing sequences. Coarse diamond milling pads address lippage, deep scratches, and binder degradation in terrazzo. Fine diamond pads (400 through 3500 grit) rebuild reflectivity in marble, limestone, and smooth slate. Diamond burnishing pads restore gloss on polished porcelain without removing material."
    },
    {
      service: "Floor Finish",
      priority: 2,
      price: "POA",
      description: "Stone and tile floor finish selection depends on material type, porosity, and environmental conditions. Marble achieves high mechanical polish through diamond abrasive sequences. Limestone and travertine reach satin or matte honed finishes. Slate receives impregnating or topical sealer for matte to gloss results. Terracotta and Victorian clay receive breathable sealers in natural, satin, or wax finishes."
    },

    // ── PRIORITY 3: Generic services worth populating ──────────────────────────

    {
      service: "Deep Cleaning",
      priority: 3,
      price: "From £25/m²",
      description: "Professional deep cleaning removes ingrained soil, failed sealers, and cleaning residues from stone and tile floors using chemistry, mechanical agitation, and full wet-vacuum extraction. Surface mopping redistributes contamination; professional extraction removes it. Chemistry is matched to surface type — alkaline for organics, targeted acid for mineral deposits on appropriate surfaces. Sealer applied on completion."
    },
    {
      service: "Routine Cleaning",
      priority: 3,
      price: "From £15/m²",
      description: "Routine professional cleaning maintains stone and tile floors between full restorations using pH-neutral chemistry and controlled mechanical agitation. Residue build-up from domestic cleaning products is removed before it accumulates into the lock-in cycle that causes progressive soiling. Grout joints are mechanically agitated. Sealer top-up applied where required."
    },
    {
      service: "Routine Maintenance",
      priority: 3,
      price: "From £15/m²",
      description: "Scheduled professional maintenance extends the life of restored stone and tile floors by addressing soil build-up, sealer wear, and grout deterioration before they require full restoration. Annual or biennial professional visits maintain the surface in its restored condition. pH-neutral cleaning guidance is provided for correct homeowner maintenance between visits."
    },
    {
      service: "Maintenance Contracts",
      priority: 3,
      price: "POA",
      description: "Scheduled maintenance agreements for commercial and residential stone and tile floors provide regular professional cleaning, sealer top-up, and grout inspection on an agreed frequency. Maintenance contracts extend restoration lifespan and reduce the cost and disruption of periodic full restorations. Assessment of current floor condition establishes the appropriate maintenance schedule."
    },
    {
      service: "Commercial Cleaning",
      priority: 3,
      price: "POA",
      description: "Commercial stone and tile floor cleaning uses professional-grade chemistry and machinery to restore high-traffic surfaces in retail, hospitality, and office environments. Assessment establishes material type, contamination depth, and sealer condition before work begins. The correct protocol is determined by what the floor is made of, not by cleaning frequency."
    },
    {
      service: "Pressure Cleaning",
      priority: 3,
      price: "POA",
      description: "Controlled pressure cleaning is used on appropriate external stone surfaces including sandstone, granite, and engineering brick. Pressure settings are matched to the material — porous sandstone requires lower pressure than dense granite to avoid surface erosion. Interior stone and historic clay tile installations are not pressure washed — controlled extraction methods are used instead."
    },
    {
      service: "Pressure Washing",
      priority: 3,
      price: "POA",
      description: "External stone pressure washing removes biological growth, dirt, and surface contamination from appropriate surfaces. Chemistry is matched to the stone type — calcite-cemented sandstone prohibits acidic treatments. Internal stone and historic tile floors are not pressure washed. Assessment before treatment confirms surface type and appropriate pressure and chemistry combination."
    },
    {
      service: "High Pressure Cleaning",
      priority: 3,
      price: "POA",
      description: "High-pressure cleaning is appropriate for durable external stone surfaces with confirmed stability. Pressure and nozzle selection are matched to the material — inappropriate settings erode softer stone surfaces and damage historic pointing. Assessment confirms suitability before work begins. Internal stone floors use mechanical agitation and wet-vacuum extraction rather than pressure washing."
    },
    {
      service: "Steam Cleaning",
      priority: 3,
      price: "POA",
      description: "Professional steam extraction is used for biological contamination in cementitious grout on ceramic and porcelain tile installations. Steam is not used on sealed natural stone floors — heat degrades topical sealers and forces moisture past impregnating systems. Domestic steam mops are not appropriate for any sealed stone floor. Professional pressurised hot water extraction with immediate removal is the correct technique."
    },
    {
      service: "Paving Clean",
      priority: 3,
      price: "POA",
      description: "External natural stone paving cleaning removes biological growth, soiling, and surface deposits from sandstone, limestone, granite, and slate paving using chemistry appropriate to the stone type. Calcite-cemented sandstone prohibits acidic treatments. Biological cleaner addresses algae and moss growth. Breathable impregnating sealer applied to sandstone and limestone paving to improve cleanability and slow re-soiling."
    },
    {
      service: "Water Damage",
      priority: 3,
      price: "POA",
      description: "Water damage assessment identifies the source and extent of moisture impact on stone and tile floors before restoration work begins. Ongoing moisture movement on floors without a damp-proof membrane requires breathable sealing systems. Efflorescence indicates active salt migration from below. Sealer failure from trapped moisture requires full stripping before correct re-sealing can be carried out."
    }

  ];

  // ── CREATE SPREADSHEET ─────────────────────────────────────────────────────
  var ss = SpreadsheetApp.create("Abbey Floor Care — GBP Service Descriptions");
  var sheet = ss.getActiveSheet();
  sheet.setName("GBP Services");

  // ── HEADERS ────────────────────────────────────────────────────────────────
  var headers = ["Priority", "Service", "Price", "Service Description", "Character Count", "Status"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ── HEADER FORMATTING ──────────────────────────────────────────────────────
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1a237e");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(11);
  sheet.setFrozenRows(1);

  // ── DATA ROWS ──────────────────────────────────────────────────────────────
  var rows = [];
  services.forEach(function(s) {
    var descLen = s.description.length;
    var status  = descLen <= 300 ? "✔ Within limit" : "⚠ Over 300 chars — trim needed";
    rows.push([
      s.priority,
      s.service,
      s.price,
      s.description,
      descLen,
      status
    ]);
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // ── COLUMN WIDTHS ──────────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 70);   // Priority
  sheet.setColumnWidth(2, 220);  // Service
  sheet.setColumnWidth(3, 120);  // Price
  sheet.setColumnWidth(4, 600);  // Description
  sheet.setColumnWidth(5, 120);  // Char count
  sheet.setColumnWidth(6, 160);  // Status

  // ── ROW FORMATTING ─────────────────────────────────────────────────────────
  // Alternate row shading by priority
  for (var i = 0; i < rows.length; i++) {
    var rowNum  = i + 2;
    var priority = rows[i][0];
    var bg = priority === 1 ? "#e8f5e9"   // green tint — priority 1
           : priority === 2 ? "#fff8e1"   // amber tint — priority 2
           : "#fce4ec";                   // pink tint  — priority 3
    sheet.getRange(rowNum, 1, 1, headers.length).setBackground(bg);
  }

  // ── WRAP TEXT IN DESCRIPTION COLUMN ───────────────────────────────────────
  sheet.getRange(2, 4, rows.length, 1).setWrap(true);

  // ── CENTRE ALIGN PRIORITY AND CHAR COUNT ──────────────────────────────────
  sheet.getRange(2, 1, rows.length, 1).setHorizontalAlignment("center");
  sheet.getRange(2, 5, rows.length, 1).setHorizontalAlignment("center");
  sheet.getRange(2, 6, rows.length, 1).setHorizontalAlignment("center");

  // ── PRIORITY LEGEND ───────────────────────────────────────────────────────
  var legendRow = rows.length + 3;
  sheet.getRange(legendRow, 1).setValue("LEGEND");
  sheet.getRange(legendRow, 1).setFontWeight("bold");
  sheet.getRange(legendRow + 1, 1, 1, 2).setValues([["1 (green)", "Named stone/tile — populate first"]]);
  sheet.getRange(legendRow + 2, 1, 1, 2).setValues([["2 (amber)", "Repair/treatment — populate next"]]);
  sheet.getRange(legendRow + 3, 1, 1, 2).setValues([["3 (pink)",  "Generic service — populate last"]]);
  sheet.getRange(legendRow + 1, 1).setBackground("#e8f5e9");
  sheet.getRange(legendRow + 2, 1).setBackground("#fff8e1");
  sheet.getRange(legendRow + 3, 1).setBackground("#fce4ec");

  // ── SUMMARY ───────────────────────────────────────────────────────────────
  var p1count = services.filter(function(s){ return s.priority === 1; }).length;
  var p2count = services.filter(function(s){ return s.priority === 2; }).length;
  var p3count = services.filter(function(s){ return s.priority === 3; }).length;
  var overLimit = rows.filter(function(r){ return r[4] > 300; }).length;

  var summaryRow = legendRow + 5;
  sheet.getRange(summaryRow, 1).setValue("SUMMARY");
  sheet.getRange(summaryRow, 1).setFontWeight("bold");
  sheet.getRange(summaryRow + 1, 1, 4, 2).setValues([
    ["Total services", rows.length],
    ["Priority 1 (stone/tile)", p1count],
    ["Priority 2 (repair/treatment)", p2count],
    ["Priority 3 (generic)", p3count],
  ]);

  SpreadsheetApp.getUi().alert(
    "GBP Service Description Sheet created.\n\n" +
    "Total services: " + rows.length + "\n" +
    "Priority 1: " + p1count + "\n" +
    "Priority 2: " + p2count + "\n" +
    "Priority 3: " + p3count + "\n\n" +
    (overLimit > 0
      ? "⚠ " + overLimit + " description(s) exceed 300 characters — see Status column."
      : "✔ All descriptions within 300 character limit.")
  );

  return ss.getUrl();
}