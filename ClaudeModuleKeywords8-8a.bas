Attribute VB_Name = "ModuleKeywords"
Option Explicit

' Safe method to add keywords without duplicates
Private Sub SafeAdd(dict As Object, key As String, value As String)
    key = LCase(Trim(key))
    value = LCase(Trim(value))
    
    ' Only add if key doesn't already exist and both key and value are not empty
    If key <> "" And value <> "" Then
        If Not dict.Exists(key) Then
            dict.Add key, value
        Else
            ' Optional: Debug duplicate keys (uncomment to see in Immediate Window)
            ' Debug.Print "Duplicate key skipped: " & key & " | Existing: " & dict(key) & " | New: " & value
        End If
    End If
End Sub

' Populate dictionary with all keywords from your CSV, cleaned and deduplicated
Public Sub PopulateKeywordDictionary(dict As Object)
    dict.RemoveAll
    
    ' Process multi-keyword entries by splitting on commas
    Dim multiKeywords As Variant
    Dim kw As Variant
    
    ' Address keywords
    SafeAdd dict, "address", "address"
    SafeAdd dict, "ave", "address"
    SafeAdd dict, "cir", "address"
    SafeAdd dict, "circle", "address"
    SafeAdd dict, "city", "address"
    SafeAdd dict, "county", "address"
    SafeAdd dict, "court", "address"
    SafeAdd dict, "ct", "address"
    SafeAdd dict, "lane", "address"
    SafeAdd dict, "ln", "address"
    SafeAdd dict, "pl", "address"
    SafeAdd dict, "rd", "address"
    SafeAdd dict, "road", "address"
    SafeAdd dict, "st", "address"
    SafeAdd dict, "street", "address"
    SafeAdd dict, "usps", "address"
    SafeAdd dict, "zip code", "address"
    
    ' Amenities
    SafeAdd dict, "pool", "amenities"
    
    ' Appraiser Details and Credentials (handling multi-keywords)
    SafeAdd dict, "appraiser qualifications", "appraiser details and credentials"
    ' Split multi-keyword entry "appraiser, credential, license"
    multiKeywords = Split("appraiser,credential,license", ",")
    For Each kw In multiKeywords
        SafeAdd dict, Trim(kw), "appraiser details and credentials"
    Next kw
    SafeAdd dict, "certifications", "appraiser details and credentials"
    SafeAdd dict, "expired e&o", "appraiser details and credentials"
    SafeAdd dict, "license numbers", "appraiser details and credentials"
    SafeAdd dict, "license", "appraiser details and credentials"
    SafeAdd dict, "e&o", "appraiser details and credentials"
    SafeAdd dict, "amc", "appraiser details and credentials"
    
    ' Assignment Type
    SafeAdd dict, "assignment type", "assignment type"
    
    ' Attic related
    SafeAdd dict, "attic", "attic related"
    
    ' Borrower and Lender Information
    SafeAdd dict, "borrower", "borrower and lender information"
    SafeAdd dict, "borrower names", "borrower and lender information"
    ' Handle multi-keyword "borrower, client"
    multiKeywords = Split("borrower,client", ",")
    For Each kw In multiKeywords
        SafeAdd dict, Trim(kw), "borrower and lender information"
    Next kw
    SafeAdd dict, "buyer", "borrower and lender information"
    SafeAdd dict, "client details", "borrower and lender information"
    SafeAdd dict, "co-borrower", "borrower and lender information"
    SafeAdd dict, "lender", "borrower and lender information"
    SafeAdd dict, "lender address", "borrower and lender information"
    SafeAdd dict, "coborrower", "borrower and lender information"
    SafeAdd dict, "borrowers", "borrower and lender information"
    SafeAdd dict, "borrower/owner", "borrower and lender information"
    SafeAdd dict, "borrower's", "borrower and lender information"
    SafeAdd dict, "borrowers'", "borrower and lender information"
    SafeAdd dict, "coborrowers", "borrower and lender information"
    
    ' Closing Cost
    SafeAdd dict, "closing cost", "closing cost"
    SafeAdd dict, "closing costs", "closing cost"
    
    ' Compliance with Standards
    SafeAdd dict, "code violations", "compliance with standards"
    multiKeywords = Split("compliance,code", ",")
    For Each kw In multiKeywords
        SafeAdd dict, Trim(kw), "compliance with standards"
    Next kw
    SafeAdd dict, "exposure time", "compliance with standards"
    SafeAdd dict, "prior services", "compliance with standards"
    SafeAdd dict, "regulations", "compliance with standards"
    SafeAdd dict, "safety standards", "compliance with standards"
    
    ' Condo Project Information
    SafeAdd dict, "commercial space", "condo project information"
    SafeAdd dict, "common space", "condo project information"
    SafeAdd dict, "condo", "condo project information"
    SafeAdd dict, "condo cert", "condo project information"
    SafeAdd dict, "condo fees", "condo project information"
    SafeAdd dict, "condo questionnaire", "condo project information"
    SafeAdd dict, "condominium", "condo project information"
    SafeAdd dict, "garden", "condo project information"
    SafeAdd dict, "high rise", "condo project information"
    SafeAdd dict, "mid-rise", "condo project information"
    SafeAdd dict, "number of units", "condo project information"
    SafeAdd dict, "project data", "condo project information"
    SafeAdd dict, "project info", "condo project information"
    SafeAdd dict, "project information", "condo project information"
    SafeAdd dict, "questionnaire", "condo project information"
    SafeAdd dict, "stories", "condo project information"
    SafeAdd dict, "to match", "condo project information"
    SafeAdd dict, "units owned", "condo project information"
    SafeAdd dict, "units rented", "condo project information"
    
    ' Contract Section
    SafeAdd dict, "addendum", "contract section"
    SafeAdd dict, "arms length", "contract section"
    SafeAdd dict, "contract date", "contract section"
    SafeAdd dict, "contract price", "contract section"
    SafeAdd dict, "contract prices", "contract section"
    SafeAdd dict, "credit", "contract section"
    SafeAdd dict, "dom", "contract section"
    SafeAdd dict, "listing history", "contract section"
    SafeAdd dict, "non arms length", "contract section"
    SafeAdd dict, "offer prices", "contract section"
    SafeAdd dict, "sale agreements", "contract section"
    SafeAdd dict, "seller credit", "contract section"
    SafeAdd dict, "spcc", "contract section"
    SafeAdd dict, "date of contract", "contract section"
    
    ' Cost Approach
    SafeAdd dict, "cost", "cost approach"
    SafeAdd dict, "depreciation", "cost approach"
    SafeAdd dict, "land values", "cost approach"
    SafeAdd dict, "rel", "cost approach"
    SafeAdd dict, "remaining economic life", "cost approach"
    SafeAdd dict, "replacement costs", "cost approach"
    
    ' Discrepancies and Inconsistencies
    SafeAdd dict, "advise", "discrepancies and inconsistencies"
    SafeAdd dict, "blank", "discrepancies and inconsistencies"
    SafeAdd dict, "blank page", "discrepancies and inconsistencies"
    SafeAdd dict, "consistency", "discrepancies and inconsistencies"
    SafeAdd dict, "differences", "discrepancies and inconsistencies"
    SafeAdd dict, "discrepancies", "discrepancies and inconsistencies"
    SafeAdd dict, "discrepancy", "discrepancies and inconsistencies"
    SafeAdd dict, "incomplete items", "discrepancies and inconsistencies"
    SafeAdd dict, "inconsistencies", "discrepancies and inconsistencies"
    SafeAdd dict, "inconsistency", "discrepancies and inconsistencies"
    SafeAdd dict, "inconsistent items", "discrepancies and inconsistencies"
    SafeAdd dict, "lender", "discrepancies and inconsistencies"
    SafeAdd dict, "lender address", "discrepancies and inconsistencies"
    SafeAdd dict, "lenders address", "discrepancies and inconsistencies"
    SafeAdd dict, "missing", "discrepancies and inconsistencies"
    SafeAdd dict, "missing items", "discrepancies and inconsistencies"
    
    ' FEMA Disaster Impact
    SafeAdd dict, "cyclone", "fema disaster impact"
    SafeAdd dict, "disaster", "fema disaster impact"
    SafeAdd dict, "disaster-related impacts", "fema disaster impact"
    SafeAdd dict, "earthquake", "fema disaster impact"
    SafeAdd dict, "environmental hazards", "fema disaster impact"
    SafeAdd dict, "fema", "fema disaster impact"
    SafeAdd dict, "fema flood zones", "fema disaster impact"
    SafeAdd dict, "fire", "fema disaster impact"
    SafeAdd dict, "fires", "fema disaster impact"
    SafeAdd dict, "flood", "fema disaster impact"
    SafeAdd dict, "flooding", "fema disaster impact"
    SafeAdd dict, "floodplains", "fema disaster impact"
    SafeAdd dict, "helene", "fema disaster impact"
    SafeAdd dict, "hurricane", "fema disaster impact"
    SafeAdd dict, "milton", "fema disaster impact"
    SafeAdd dict, "storm", "fema disaster impact"
    SafeAdd dict, "tropical storm", "fema disaster impact"
    SafeAdd dict, "water-related risks", "fema disaster impact"
    
    ' FHA/USDA Requirements
    SafeAdd dict, "4000.1", "fha/usda requirements"
    SafeAdd dict, "case number", "fha/usda requirements"
    SafeAdd dict, "fha", "fha/usda requirements"
    SafeAdd dict, "hud", "fha/usda requirements"
    SafeAdd dict, "mpr", "fha/usda requirements"
    SafeAdd dict, "mps", "fha/usda requirements"
    SafeAdd dict, "usda", "fha/usda requirements"
    SafeAdd dict, "well and septic distance", "fha/usda requirements"
    SafeAdd dict, "hud/fha", "fha/usda requirements"
    
    ' Handrail
    SafeAdd dict, "handrail", "handrail"
    SafeAdd dict, "hand rail", "handrail"
    SafeAdd dict, "handrails", "handrail"
    SafeAdd dict, "hand rails", "handrail"
    
    ' Income Approach
    SafeAdd dict, "1007", "income approach"
    SafeAdd dict, "capitalization rates", "income approach"
    SafeAdd dict, "income", "income approach"
    SafeAdd dict, "landlord", "income approach"
    SafeAdd dict, "lease", "income approach"
    SafeAdd dict, "operating expenses", "income approach"
    SafeAdd dict, "rent", "income approach"
    SafeAdd dict, "rental income analysis", "income approach"
    SafeAdd dict, "tenant", "income approach"
    
    ' Inspection Requirements
    SafeAdd dict, "builder", "inspection requirements"
    SafeAdd dict, "carbon monoxide", "inspection requirements"
    SafeAdd dict, "co", "inspection requirements"
    SafeAdd dict, "deficiencies", "inspection requirements"
    SafeAdd dict, "detectors", "inspection requirements"
    SafeAdd dict, "disaster", "inspection requirements"
    SafeAdd dict, "double strap", "inspection requirements"
    SafeAdd dict, "inspection", "inspection requirements"
    SafeAdd dict, "inspection issues", "inspection requirements"
    SafeAdd dict, "physical deficiencies", "inspection requirements"
    SafeAdd dict, "property evaluation", "inspection requirements"
    SafeAdd dict, "smoke", "inspection requirements"
    SafeAdd dict, "smoke detectors", "inspection requirements"
    SafeAdd dict, "water heater", "inspection requirements"
    SafeAdd dict, "co detect", "inspection requirements"
    
    ' Legal and Ownership Details
    SafeAdd dict, "apn", "legal and ownership details"
    SafeAdd dict, "county", "legal and ownership details"
    SafeAdd dict, "deed", "legal and ownership details"
    SafeAdd dict, "deed restriction", "legal and ownership details"
    SafeAdd dict, "descriptions", "legal and ownership details"
    SafeAdd dict, "legal", "legal and ownership details"
    SafeAdd dict, "legal definitions", "legal and ownership details"
    SafeAdd dict, "legal ownership", "legal and ownership details"
    SafeAdd dict, "occupancy", "legal and ownership details"
    SafeAdd dict, "oopr", "legal and ownership details"
    SafeAdd dict, "oor", "legal and ownership details"
    SafeAdd dict, "owner", "legal and ownership details"
    SafeAdd dict, "owner of record", "legal and ownership details"
    SafeAdd dict, "ownership", "legal and ownership details"
    SafeAdd dict, "pr", "legal and ownership details"
    SafeAdd dict, "public record", "legal and ownership details"
    SafeAdd dict, "section", "legal and ownership details"
    SafeAdd dict, "seller", "legal and ownership details"
    SafeAdd dict, "tenant", "legal and ownership details"
    SafeAdd dict, "title", "legal and ownership details"
    SafeAdd dict, "titled", "legal and ownership details"
    
    ' Map/Exhibits
    SafeAdd dict, "aerial", "map/exhibits"
    SafeAdd dict, "comp map", "map/exhibits"
    SafeAdd dict, "comparable map", "map/exhibits"
    SafeAdd dict, "exhibit", "map/exhibits"
    SafeAdd dict, "flood map", "map/exhibits"
    SafeAdd dict, "location map", "map/exhibits"
    SafeAdd dict, "maps", "map/exhibits"
    SafeAdd dict, "plat", "map/exhibits"
    SafeAdd dict, "survey", "map/exhibits"
    
    ' Market Trends
    SafeAdd dict, "change in value", "market trends"
    SafeAdd dict, "decrease", "market trends"
    SafeAdd dict, "decreasing", "market trends"
    SafeAdd dict, "home price index", "market trends"
    SafeAdd dict, "hpi", "market trends"
    SafeAdd dict, "increase", "market trends"
    SafeAdd dict, "increasing", "market trends"
    SafeAdd dict, "market trend", "market trends"
    SafeAdd dict, "market trends", "market trends"
    SafeAdd dict, "price changes", "market trends"
    SafeAdd dict, "stable market adjustment", "market trends"
    SafeAdd dict, "time adjustment", "market trends"
    SafeAdd dict, "marketability", "market trends"
    
    ' Neighborhood Description
    SafeAdd dict, "area details", "neighborhood description"
    SafeAdd dict, "boundaries", "neighborhood description"
    SafeAdd dict, "built-up", "neighborhood description"
    SafeAdd dict, "busy", "neighborhood description"
    SafeAdd dict, "community", "neighborhood description"
    SafeAdd dict, "demand", "neighborhood description"
    SafeAdd dict, "demand/supply", "neighborhood description"
    SafeAdd dict, "external factors", "neighborhood description"
    SafeAdd dict, "location", "neighborhood description"
    SafeAdd dict, "neighborhood", "neighborhood description"
    SafeAdd dict, "neighborhood name", "neighborhood description"
    SafeAdd dict, "over improved", "neighborhood description"
    SafeAdd dict, "private road", "neighborhood description"
    SafeAdd dict, "rural", "neighborhood description"
    SafeAdd dict, "suburban", "neighborhood description"
    SafeAdd dict, "supply", "neighborhood description"
    SafeAdd dict, "surroundings", "neighborhood description"
    SafeAdd dict, "under improved", "neighborhood description"
    
    ' New Construction
    SafeAdd dict, "budget", "new construction"
    SafeAdd dict, "builder", "new construction"
    SafeAdd dict, "builder name", "new construction"
    SafeAdd dict, "builders name", "new construction"
    SafeAdd dict, "builder's name", "new construction"
    SafeAdd dict, "new construction", "new construction"
    SafeAdd dict, "plans", "new construction"
    SafeAdd dict, "plans and specs", "new construction"
    SafeAdd dict, "specs", "new construction"
    
    ' Photos and Documentation
    SafeAdd dict, "album", "photos and documentation"
    SafeAdd dict, "archive", "photos and documentation"
    SafeAdd dict, "blur the people", "photos and documentation"
    SafeAdd dict, "blur the person", "photos and documentation"
    SafeAdd dict, "blur the photo", "photos and documentation"
    SafeAdd dict, "camera", "photos and documentation"
    SafeAdd dict, "capture", "photos and documentation"
    SafeAdd dict, "catalog", "photos and documentation"
    SafeAdd dict, "collection", "photos and documentation"
    SafeAdd dict, "compilation", "photos and documentation"
    SafeAdd dict, "depiction", "photos and documentation"
    SafeAdd dict, "documentation", "photos and documentation"
    SafeAdd dict, "dossier", "photos and documentation"
    SafeAdd dict, "evidence", "photos and documentation"
    SafeAdd dict, "file", "photos and documentation"
    SafeAdd dict, "folder", "photos and documentation"
    SafeAdd dict, "frame", "photos and documentation"
    SafeAdd dict, "gallery", "photos and documentation"
    SafeAdd dict, "illustration", "photos and documentation"
    SafeAdd dict, "image", "photos and documentation"
    SafeAdd dict, "imagery", "photos and documentation"
    SafeAdd dict, "images", "photos and documentation"
    SafeAdd dict, "index", "photos and documentation"
    SafeAdd dict, "inventory", "photos and documentation"
    SafeAdd dict, "log", "photos and documentation"
    SafeAdd dict, "photo", "photos and documentation"
    SafeAdd dict, "photograph", "photos and documentation"
    SafeAdd dict, "photos", "photos and documentation"
    SafeAdd dict, "pic", "photos and documentation"
    SafeAdd dict, "pictorial", "photos and documentation"
    SafeAdd dict, "picture", "photos and documentation"
    SafeAdd dict, "pictures", "photos and documentation"
    SafeAdd dict, "plat map", "photos and documentation"
    SafeAdd dict, "portfolio", "photos and documentation"
    SafeAdd dict, "portrait", "photos and documentation"
    SafeAdd dict, "proof", "photos and documentation"
    SafeAdd dict, "record", "photos and documentation"
    SafeAdd dict, "register", "photos and documentation"
    SafeAdd dict, "repository", "photos and documentation"
    SafeAdd dict, "representation", "photos and documentation"
    SafeAdd dict, "scene", "photos and documentation"
    SafeAdd dict, "shot", "photos and documentation"
    SafeAdd dict, "snapshot", "photos and documentation"
    SafeAdd dict, "still", "photos and documentation"
    SafeAdd dict, "view", "photos and documentation"
    SafeAdd dict, "visual", "photos and documentation"
    SafeAdd dict, "visual aid", "photos and documentation"
    SafeAdd dict, "photographs", "photos and documentation"
    SafeAdd dict, "photo(s)", "photos and documentation"
    
    ' Property Description
    SafeAdd dict, "boundary disputes", "property description"
    SafeAdd dict, "condition", "property description"
    SafeAdd dict, "easement", "property description"
    SafeAdd dict, "easements", "property description"
    SafeAdd dict, "encroachment", "property description"
    SafeAdd dict, "encroachments", "property description"
    SafeAdd dict, "finished", "property description"
    SafeAdd dict, "finishes", "property description"
    SafeAdd dict, "location", "property description"
    SafeAdd dict, "lot size", "property description"
    SafeAdd dict, "parcel", "property description"
    SafeAdd dict, "quality", "property description"
    SafeAdd dict, "rating", "property description"
    SafeAdd dict, "right-of-way issues", "property description"
    SafeAdd dict, "site area", "property description"
    SafeAdd dict, "site size", "property description"
    SafeAdd dict, "structure", "property description"
    SafeAdd dict, "view", "property description"
    SafeAdd dict, "year built", "property description"
    
    ' Property Taxes and Assessments
    SafeAdd dict, "assessed values", "property taxes and assessments"
    SafeAdd dict, "assessor", "property taxes and assessments"
    SafeAdd dict, "estimated taxes", "property taxes and assessments"
    SafeAdd dict, "property tax rates", "property taxes and assessments"
    SafeAdd dict, "real estate taxes", "property taxes and assessments"
    SafeAdd dict, "special assessment", "property taxes and assessments"
    SafeAdd dict, "tax assessment", "property taxes and assessments"
    SafeAdd dict, "tax assessments", "property taxes and assessments"
    SafeAdd dict, "taxes", "property taxes and assessments"
    
    ' PUD/HOA Information
    SafeAdd dict, "annual", "pud/hoa information"
    SafeAdd dict, "hoa", "pud/hoa information"
    SafeAdd dict, "hoa fee", "pud/hoa information"
    SafeAdd dict, "hoa fees", "pud/hoa information"
    SafeAdd dict, "hoa-related information", "pud/hoa information"
    SafeAdd dict, "homeowners association", "pud/hoa information"
    SafeAdd dict, "monthly", "pud/hoa information"
    SafeAdd dict, "planned unit", "pud/hoa information"
    SafeAdd dict, "pud", "pud/hoa information"
    SafeAdd dict, "pud details", "pud/hoa information"
    
    ' Purchase Price
    SafeAdd dict, "purchase price", "purchase price"
    SafeAdd dict, "sale price", "purchase price"
    SafeAdd dict, "sales price", "purchase price"
    
    ' Re-built
    SafeAdd dict, "can subject be re-built to its current footprint if destroyed", "re-built"
    
    ' Repairs and Maintenance
    SafeAdd dict, "asbestos", "repairs and maintenance"
    SafeAdd dict, "broken", "repairs and maintenance"
    SafeAdd dict, "contamination", "repairs and maintenance"
    SafeAdd dict, "cost to cure", "repairs and maintenance"
    SafeAdd dict, "ctc", "repairs and maintenance"
    SafeAdd dict, "damage", "repairs and maintenance"
    SafeAdd dict, "damaged", "repairs and maintenance"
    SafeAdd dict, "deferred", "repairs and maintenance"
    SafeAdd dict, "environmental", "repairs and maintenance"
    SafeAdd dict, "hazard", "repairs and maintenance"
    SafeAdd dict, "hazardous materials", "repairs and maintenance"
    SafeAdd dict, "improvements", "repairs and maintenance"
    SafeAdd dict, "lead", "repairs and maintenance"
    SafeAdd dict, "lead paint", "repairs and maintenance"
    SafeAdd dict, "maintain", "repairs and maintenance"
    SafeAdd dict, "maintained", "repairs and maintenance"
    SafeAdd dict, "maintenance", "repairs and maintenance"
    SafeAdd dict, "moisture", "repairs and maintenance"
    SafeAdd dict, "mold", "repairs and maintenance"
    SafeAdd dict, "repair", "repairs and maintenance"
    SafeAdd dict, "repaired", "repairs and maintenance"
    SafeAdd dict, "repairs", "repairs and maintenance"
    SafeAdd dict, "sor", "repairs and maintenance"
    SafeAdd dict, "sow", "repairs and maintenance"
    SafeAdd dict, "toxic", "repairs and maintenance"
    SafeAdd dict, "water", "repairs and maintenance"
    SafeAdd dict, "workmanlike", "repairs and maintenance"
    SafeAdd dict, "workman-like", "repairs and maintenance"
    
    ' Roof Related
    SafeAdd dict, "roof", "roof related"
    
    ' Sales Comparison Approach/Adjustments
    SafeAdd dict, "above grade", "sales comparison approach/adjustments"
    SafeAdd dict, "access", "sales comparison approach/adjustments"
    SafeAdd dict, "adjusted", "sales comparison approach/adjustments"
    SafeAdd dict, "adjustment", "sales comparison approach/adjustments"
    SafeAdd dict, "adjustment grids", "sales comparison approach/adjustments"
    SafeAdd dict, "below grade", "sales comparison approach/adjustments"
    SafeAdd dict, "bracket", "sales comparison approach/adjustments"
    SafeAdd dict, "bracketed", "sales comparison approach/adjustments"
    SafeAdd dict, "bracketing", "sales comparison approach/adjustments"
    SafeAdd dict, "comp", "sales comparison approach/adjustments"
    SafeAdd dict, "comparable", "sales comparison approach/adjustments"
    SafeAdd dict, "comparable properties", "sales comparison approach/adjustments"
    SafeAdd dict, "comparable sale", "sales comparison approach/adjustments"
    SafeAdd dict, "comparable sales", "sales comparison approach/adjustments"
    SafeAdd dict, "distance", "sales comparison approach/adjustments"
    SafeAdd dict, "gross adjustment", "sales comparison approach/adjustments"
    SafeAdd dict, "gross net line", "sales comparison approach/adjustments"
    SafeAdd dict, "line item excessive", "sales comparison approach/adjustments"
    SafeAdd dict, "market analysis", "sales comparison approach/adjustments"
    SafeAdd dict, "mile", "sales comparison approach/adjustments"
    SafeAdd dict, "narrative", "sales comparison approach/adjustments"
    SafeAdd dict, "net adjustment", "sales comparison approach/adjustments"
    SafeAdd dict, "parameter", "sales comparison approach/adjustments"
    SafeAdd dict, "parameters", "sales comparison approach/adjustments"
    SafeAdd dict, "proximate", "sales comparison approach/adjustments"
    SafeAdd dict, "proximity", "sales comparison approach/adjustments"
    SafeAdd dict, "sales", "sales comparison approach/adjustments"
    SafeAdd dict, "search", "sales comparison approach/adjustments"
    SafeAdd dict, "un bracketed", "sales comparison approach/adjustments"
    SafeAdd dict, "unadjusted", "sales comparison approach/adjustments"
    SafeAdd dict, "unsupported", "sales comparison approach/adjustments"
    SafeAdd dict, "comps", "sales comparison approach/adjustments"
    SafeAdd dict, "comparables", "sales comparison approach/adjustments"
    SafeAdd dict, "comparable(s)", "sales comparison approach/adjustments"
    SafeAdd dict, "adjustments", "sales comparison approach/adjustments"
    SafeAdd dict, "adjustment(s)", "sales comparison approach/adjustments"
    
    ' Signature
    SafeAdd dict, "signature", "signature"
    
    ' Sketch and Floor Plan Issues
    SafeAdd dict, "ansi", "sketch and floor plan issues"
    SafeAdd dict, "architectural drawing", "sketch and floor plan issues"
    SafeAdd dict, "arrangement", "sketch and floor plan issues"
    SafeAdd dict, "blueprint", "sketch and floor plan issues"
    SafeAdd dict, "configuration", "sketch and floor plan issues"
    SafeAdd dict, "depiction", "sketch and floor plan issues"
    SafeAdd dict, "design", "sketch and floor plan issues"
    SafeAdd dict, "diagram", "sketch and floor plan issues"
    SafeAdd dict, "floor plan", "sketch and floor plan issues"
    SafeAdd dict, "floorplan", "sketch and floor plan issues"
    SafeAdd dict, "gla", "sketch and floor plan issues"
    SafeAdd dict, "gross living area", "sketch and floor plan issues"
    SafeAdd dict, "label", "sketch and floor plan issues"
    SafeAdd dict, "layout", "sketch and floor plan issues"
    SafeAdd dict, "layout problem", "sketch and floor plan issues"
    SafeAdd dict, "layout problems", "sketch and floor plan issues"
    SafeAdd dict, "outline", "sketch and floor plan issues"
    SafeAdd dict, "plan", "sketch and floor plan issues"
    SafeAdd dict, "property dimensions", "sketch and floor plan issues"
    SafeAdd dict, "rendering", "sketch and floor plan issues"
    SafeAdd dict, "schematic", "sketch and floor plan issues"
    SafeAdd dict, "scheme", "sketch and floor plan issues"
    SafeAdd dict, "sketch", "sketch and floor plan issues"
    
    ' Subject To
    SafeAdd dict, "subject to", "subject to"
    SafeAdd dict, "subject to repairs", "subject to"
    
    ' Tax Doc
    SafeAdd dict, "tax", "tax doc"
    
    ' Utility Information
    SafeAdd dict, "connection", "utility information"
    SafeAdd dict, "covenant", "utility information"
    SafeAdd dict, "electricity", "utility information"
    SafeAdd dict, "gas services", "utility information"
    SafeAdd dict, "maintenance agreement", "utility information"
    SafeAdd dict, "oil", "utility information"
    SafeAdd dict, "private", "utility information"
    SafeAdd dict, "private road", "utility information"
    SafeAdd dict, "private street", "utility information"
    SafeAdd dict, "private well", "utility information"
    SafeAdd dict, "propane", "utility information"
    SafeAdd dict, "public utilities", "utility information"
    SafeAdd dict, "road", "utility information"
    SafeAdd dict, "septic", "utility information"
    SafeAdd dict, "sewer", "utility information"
    SafeAdd dict, "solar", "utility information"
    SafeAdd dict, "solar panels", "utility information"
    SafeAdd dict, "street", "utility information"
    SafeAdd dict, "utility", "utility information"
    SafeAdd dict, "water", "utility information"
    SafeAdd dict, "water supply", "utility information"
    SafeAdd dict, "wells", "utility information"
    SafeAdd dict, "utilities", "utility information"
    
    ' Value Conclusion
    SafeAdd dict, "reconciliation", "value conclusion"
    SafeAdd dict, "value", "value conclusion"
    SafeAdd dict, "value conclusion", "value conclusion"
    
    ' Zoning and Highest and Best Use (HBU)
    SafeAdd dict, "55+", "zoning and highest and best use (hbu)"
    SafeAdd dict, "accessory unit", "zoning and highest and best use (hbu)"
    SafeAdd dict, "adu", "zoning and highest and best use (hbu)"
    SafeAdd dict, "age restricted", "zoning and highest and best use (hbu)"
    SafeAdd dict, "coach house", "zoning and highest and best use (hbu)"
    SafeAdd dict, "commercial", "zoning and highest and best use (hbu)"
    SafeAdd dict, "conform", "zoning and highest and best use (hbu)"
    SafeAdd dict, "conforming", "zoning and highest and best use (hbu)"
    SafeAdd dict, "conforms", "zoning and highest and best use (hbu)"
    SafeAdd dict, "farm", "zoning and highest and best use (hbu)"
    SafeAdd dict, "guest house", "zoning and highest and best use (hbu)"
    SafeAdd dict, "hbu", "zoning and highest and best use (hbu)"
    SafeAdd dict, "highest and best use", "zoning and highest and best use (hbu)"
    SafeAdd dict, "illegal", "zoning and highest and best use (hbu)"
    SafeAdd dict, "in law", "zoning and highest and best use (hbu)"
    SafeAdd dict, "in-law", "zoning and highest and best use (hbu)"
    SafeAdd dict, "land use", "zoning and highest and best use (hbu)"
    SafeAdd dict, "land use restrictions", "zoning and highest and best use (hbu)"
    SafeAdd dict, "legality", "zoning and highest and best use (hbu)"
    SafeAdd dict, "nonconfirming", "zoning and highest and best use (hbu)"
    SafeAdd dict, "non-conforming", "zoning and highest and best use (hbu)"
    SafeAdd dict, "permit", "zoning and highest and best use (hbu)"
    SafeAdd dict, "permits", "zoning and highest and best use (hbu)"
    SafeAdd dict, "permitted", "zoning and highest and best use (hbu)"
    SafeAdd dict, "permitted uses", "zoning and highest and best use (hbu)"
    SafeAdd dict, "permitting", "zoning and highest and best use (hbu)"
    SafeAdd dict, "related living", "zoning and highest and best use (hbu)"
    SafeAdd dict, "senior", "zoning and highest and best use (hbu)"
    SafeAdd dict, "survey", "zoning and highest and best use (hbu)"
    SafeAdd dict, "zone", "zoning and highest and best use (hbu)"
    SafeAdd dict, "zoned", "zoning and highest and best use (hbu)"
    SafeAdd dict, "zoning", "zoning and highest and best use (hbu)"
    SafeAdd dict, "zoning classifications", "zoning and highest and best use (hbu)"
    SafeAdd dict, "zoning land use", "zoning and highest and best use (hbu)"
    
End Sub

' Optional: Debug function to check for duplicates and report statistics
Public Sub DebugKeywordDictionary()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    PopulateKeywordDictionary dict
    
    Debug.Print "=== KEYWORD DICTIONARY DEBUG REPORT ==="
    Debug.Print "Total unique keywords: " & dict.Count
    
    ' Show sample of categories and their counts
    Dim categoryCount As Object
    Set categoryCount = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant
    For Each key In dict.Keys
        Dim category As String
        category = dict(key)
        If categoryCount.Exists(category) Then
            categoryCount(category) = categoryCount(category) + 1
        Else
            categoryCount.Add category, 1
        End If
    Next key
    
    Debug.Print "=== CATEGORIES AND KEYWORD COUNTS ==="
    Dim cat As Variant
    For Each cat In categoryCount.Keys
        Debug.Print cat & ": " & categoryCount(cat) & " keywords"
    Next cat
    
    Debug.Print "=== DEBUG COMPLETE ==="
End Sub
