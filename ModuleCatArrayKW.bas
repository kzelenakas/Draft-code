Attribute VB_Name = "ModuleKeywords"
Option Explicit

' Safely adds a key to the dictionary, skipping duplicates
Private Sub SafeAdd(dict As Object, key As String, value As String)
    key = LCase(Trim(key))
    value = LCase(Trim(value))
    If key <> "" And value <> "" Then
        If Not dict.Exists(key) Then
            dict.Add key, value
        Else
            ' Optional: Uncomment to debug duplicates
            ' If dict(key) <> value Then
            '     Debug.Print "Duplicate key with different value: " & key & _
            '                 " -> Existing: " & dict(key) & " | New: " & value
            ' End If
        End If
    End If
End Sub

' Populate the dictionary with cleaned and deduplicated keywords and categories
Public Sub PopulateKeywordDictionary(dict As Object)
    Dim items As Variant
    Dim i As Long, kwArray As Variant, kw As Variant

    ' Clear dictionary content first
    dict.RemoveAll

    ' Define all keyword-category pairs here
    ' Multi-keywords are comma-separated in keyword field, each is added separately
    items = Array( _
        Array("address", "address"), _
        Array("ave", "address"), _
        Array("cir", "address"), _
        Array("circle", "address"), _
        Array("city", "address"), _
        Array("county", "address"), _
        Array("court", "address"), _
        Array("ct", "address"), _
        Array("lane", "address"), _
        Array("ln", "address"), _
        Array("pl", "address"), _
        Array("rd", "address"), _
        Array("road", "address"), _
        Array("st", "address"), _
        Array("street", "address"), _
        Array("usps", "address"), _
        Array("zip code", "address"), _
        Array("pool", "ammenities"), _
        Array("appraiser, credential, license", "appraiser details"), _
        Array("certifications", "appraiser details"), _
        Array("expired e&o", "appraiser details"), _
        Array("license numbers", "appraiser details"), _
        Array("e&o", "appraiser details"), _
        Array("amc", "appraiser details"), _
        Array("assignment type", "assignment type"), _
        Array("attic", "attic related"), _
        Array("borrower, client", "borrower information"), _
        Array("borrower names", "borrower information"), _
        Array("buyer", "borrower information"), _
        Array("co-borrower, coborrower", "borrower information"), _
        Array("borrowers, borrower/owner", "borrower information"), _
        Array("closing cost, closing costs", "closing cost"), _
        Array("code violations, compliance code", "compliance"), _
        Array("exposure time", "compliance"), _
        Array("prior services", "compliance"), _
        Array("regulations", "compliance"), _
        Array("safety standards", "compliance"), _
        Array("commercial space, garden, high rise, mid-rise", "condo project"), _
        Array("condo, condo cert, condo questionnaire", "condo project"), _
        Array("number of units, project data, project info, project information", "condo project"), _
        Array("questionnaire, stories, units owned, units rented", "condo project"), _
        Array("addendum", "contract section"), _
        Array("arms length, non arms length", "contract section"), _
        Array("contract date, contract price, contract prices", "contract section"), _
        Array("credit, dom, listing history, offer prices, sale agreements", "contract section"), _
        Array("seller credit", "contract section"), _
        Array("spcc", "contract section"), _
        Array("cost, depreciation, land values, rel, remaining economic life", "cost approach"), _
        Array("advise, blank, blank page, consistency, differences, discrepancies", "discrepancies"), _
        Array("incomplete items, inconsistencies, inconsistency, inconsistent items", "discrepancies"), _
        Array("lender, lender address, lenders address", "discrepancies"), _
        Array("missing, missing items", "discrepancies"), _
        Array("cyclone, disaster, disaster-related impacts", "fema disaster"), _
        Array("earthquake, environmental hazards, fema", "fema disaster"), _
        Array("fema flood zones, flood, flooding, floodplains", "fema disaster"), _
        Array("helene, hurricane, milton, storm, tropical storm", "fema disaster"), _
        Array("water-related risks", "fema disaster"), _
        Array("4000.1, fha, hud, mpr, mps, usda", "fha/usda"), _
        Array("handrail, hand rails, hand rail, handrails", "handrail"), _
        Array("1007, capitalization rates, income, landlord, lease, operating expenses, rent", "income approach"), _
        Array("rental income analysis, tenant", "income approach"), _
        Array("builder, carbon monoxide, co, deficiencies, detectors, double strap", "inspection"), _
        Array("inspection, inspection issues, physical deficiencies, property evaluation", "inspection"), _
        Array("smoke, smoke detectors, water heater, co detect", "inspection"), _
        Array("apn, county, deed, deed restriction, descriptions, legal, legal definitions", "legal"), _
        Array("legal ownership, occupancy, oopr, oor, owner, owner of record, ownership", "legal"), _
        Array("pr, public record, section, seller, tenant, title, titled", "legal"), _
        Array("aerial, comp map, comparable map, exhibit, flood map, location map, maps, plat, survey", "maps"), _
        Array("change in value, decrease, home price index, hpi, increase", "market trends"), _
        Array("market trend, market trends, price changes, stable market adjustment, time adjustment", "market trends"), _
        Array("marketability", "market trends"), _
        Array("area, boundaries, built-up, busy, community, demand, demand/supply", "neighborhood"), _
        Array("external factors, location, neighborhood, neighborhood name, over improved", "neighborhood"), _
        Array("private road, rural, suburban, supply, surroundings, under improved", "neighborhood"), _
        Array("budget, builder, builder name, builders name, builder's name, new construction", "new construction"), _
        Array("plans, plans and specs, specs", "new construction"), _
        Array("album, archive, blur the people, blur the person, blur the photo", "photos"), _
        Array("camera, capture, catalog, collection, compilation, depiction, documentation", "photos"), _
        Array("dossier, evidence, file, folder, frame, gallery, illustration", "photos"), _
        Array("image, imagery, images, index, inventory, log, photo, photograph", "photos"), _
        Array("photos, pic, pictorial, picture, pictures, plat map, portfolio, portrait", "photos"), _
        Array("proof, record, register, repository, representation, scene, shot, snapshot", "photos"), _
        Array("still, view, visual, visual aid", "photos"), _
        Array("boundary disputes, condition, easement, easements, encroachment, encroachments", "property"), _
        Array("finished, finishes, location, lot size, parcel, quality, rating, right-of-way", "property"), _
        Array("site area, site size, structure, view, year built", "property"), _
        Array("assessed values, assessor, estimated taxes, property tax rates, real estate taxes", "taxes"), _
        Array("special assessment, tax assessment, tax assessments, taxes", "taxes"), _
        Array("annual, hoa, hoa fee, hoa fees, hoa-related information, homeowners association", "hoa"), _
        Array("monthly, planned unit, pud, pud details", "hoa"), _
        Array("purchase price, sale price, sales price", "purchase price"), _
        Array("asbestos, broken, contamination, cost to cure, ctc", "repairs"), _
        Array("damage, damaged, deferred, environmental, hazard, hazardous materials", "repairs"), _
        Array("improvements, lead, lead paint, maintain, maintained, maintenance", "repairs"), _
        Array("moisture, mold, repair, repaired, repairs, sor, sow, toxic", "repairs"), _
        Array("water, workmanlike, workman-like", "repairs"), _
        Array("roof", "roof"), _
        Array("above grade, access, adjusted, adjustment", "sales approach"), _
        Array("adjustment grids, below grade, bracket, bracketed, bracketing", "sales approach"), _
        Array("comp, comparable, comparable properties, comparable sale, comparable sales", "sales approach"), _
        Array("distance, gross adjustment, gross net line", "sales approach"), _
        Array("line item excessive, market analysis, mile, narrative", "sales approach"), _
        Array("net adjustment, parameter, parameters, proximate, proximity", "sales approach"), _
        Array("sales, search, un bracketed, unadjusted, unsupported", "sales approach"), _
        Array("comps, comparables, adjustments", "sales approach"), _
        Array("signature", "signature"), _
        Array("ansi, architectural drawing, arrangement, blueprint", "sketch"), _
        Array("configuration, depiction, design, diagram", "sketch"), _
        Array("floor plan, floorplan, gla, gross living area", "sketch"), _
        Array("label, layout, layout problem, layout problems", "sketch"), _
        Array("outline, plan, property dimensions, rendering, schematic, scheme, sketch", "sketch"), _
        Array("subject to", "subject to"), _
        Array("tax", "tax doc"), _
        Array("connection, covenant, electricity, gas services", "utility"), _
        Array("maintenance agreement, oil, private, private road", "utility"), _
        Array("private street, private well, propane, public utilities", "utility"), _
        Array("road, septic, sewer, solar, solar panels", "utility"), _
        Array("street, utility, water, water supply, wells, utilities", "utility"), _
        Array("reconciliation, value, value conclusion", "value conclusion"), _
        Array("55+, accessory unit, adu, age restricted", "zoning"), _
        Array("coach house, commercial, conform, conforming, conforms", "zoning"), _
        Array("farm, guest house, hbu, highest and best use", "zoning"), _
        Array("illegal, in law, in-law, land use, land use restrictions", "zoning"), _
        Array("legality, nonconfirming, non-conforming, permit, permits", "zoning"), _
        Array("permitted, permitted uses, permitting, related living", "zoning"), _
        Array("senior, survey, zone, zoned, zoning", "zoning"), _
        Array("pud, hoa", "hoa"), _
        Array("date of contract", "contract section") _
    )

    For i = LBound(items) To UBound(items)
        kwArray = Split(items(i)(0), ",")
        For Each kw In kwArray
            SafeAdd dict, kw, items(i)(1)
        Next kw
    Next i
End Sub

' SafeAdd Sub included above to avoid duplicates

' The main categorization code should call 'PopulateKeywordDictionary' and process the data accordingly.
