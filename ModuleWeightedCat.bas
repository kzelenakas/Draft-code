Attribute VB_Name = "Module1"
Option Explicit

' Type definition for Category structure
Type Category
    Name As String
    keywords() As String
    weights() As Single
End Type

Sub CategorizeTextWithWeight()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim cell As Range
    Dim lastRow As Long
    Dim uniqueCategories As Collection
    Dim categoryDict As Object
    Dim cats() As Category
    Dim i As Integer, j As Integer
    Dim categoryName As String
    Dim tempKeywords As Collection
    Dim tempWeights As Collection
    Dim totalCount As Long, categorizedCount As Long
    
    Set ws = ThisWorkbook.Worksheets("pivot")
    
    ' Validate worksheet
    If ws Is Nothing Then
        MsgBox "Pivot worksheet not found.", vbCritical
        Exit Sub
    End If
    
    ' Find last row in column G
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' Set header for output column L
    ws.Range("L1").value = "Weighted Theme"
    If lastRow >= 2 Then ws.Range("L2:L" & lastRow).ClearContents
    
    ' Get unique categories and create dictionary for deduplication
    Set uniqueCategories = GetUniqueCategories()
    Set categoryDict = CreateObject("Scripting.Dictionary")
    
    ' Create dynamic array based on actual categories
    ReDim cats(1 To uniqueCategories.Count)
    
    ' Populate categories with deduplicated keywords from CSV data
    For i = 1 To uniqueCategories.Count
        categoryName = uniqueCategories(i)
        cats(i).Name = categoryName
        
        ' Initialize collections for deduplication
        Set tempKeywords = New Collection
        Set tempWeights = New Collection
        
        ' Get keywords for this category with weights based on CSV data
        Select Case LCase(categoryName)
            Case "address"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("address", "ave", "cir", "circle", "city", "county", "court", "ct", "lane", "ln", "pl", "rd", "road", "st", "street", "usps", "zip code"), _
                Array(2.5, 1.5, 1, 1.5, 2, 1.5, 1, 1, 1, 1, 1, 1, 1.5, 1.5, 1.5, 1, 2))
                
            Case "ammenities"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("pool"), Array(2))
                
            Case "appraiser details and credentials"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("appraiser qualifications", "appraiser", "credential", "license", "certifications", "expired e&o", "license numbers", "e&o", "amc"), _
                Array(2.5, 2.5, 2, 2.5, 2, 2, 2, 2, 1.5))
                
            Case "assignment type"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("assignment type"), Array(2.5))
                
            Case "attic related"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("attic"), Array(2))
                
            Case "borrower and lender information"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("borrower", "borrower names", "client", "buyer", "client details", "co-borrower", "lender", "lender address", "coborrower", "borrowers"), _
                Array(2.5, 2, 2, 2, 2, 2, 2.5, 2, 2, 2))
                
            Case "closing cost"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("closing cost", "closing costs"), Array(2.5, 2.5))
                
            Case "compliance with standards"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("code violations", "compliance", "code", "exposure time", "prior services", "regulations", "safety standards"), _
                Array(2, 2, 2, 2, 2, 2, 2))
                
            Case "condo project information"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("commercial space", "common space", "condo", "condo cert", "condo fees", "condo questionnaire", "condominium", "garden", "high rise", "mid-rise", "number of units", "project data", "project info", "project information", "questionnaire", "stories", "units owned", "units rented"), _
                Array(1.5, 2, 2.5, 2, 2, 2, 2.5, 1, 1.5, 1.5, 2, 2, 2, 2, 2, 1.5, 2, 2))
                
            Case "contract section"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("addendum", "arms length", "contract date", "contract price", "contract prices", "credit", "dom", "listing history", "non arms length", "offer prices", "sale agreements", "seller credit", "spcc", "date of contract"), _
                Array(2, 1.5, 2.5, 2.5, 2.5, 1.5, 2, 1.5, 1.5, 2, 2, 2, 2, 2.5))
                
            Case "cost approach"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("cost", "depreciation", "land values", "rel", "remaining economic life", "replacement costs"), _
                Array(2, 2.5, 2, 2, 2.5, 2))
                
            Case "discrepancies and inconsistencies"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("advise", "blank", "blank page", "consistency", "differences", "discrepancies", "discrepancy", "incomplete items", "inconsistencies", "inconsistency", "inconsistent items", "lender", "lender address", "lenders address", "missing", "missing items"), _
                Array(1.5, 2, 2, 2, 2, 2.5, 2.5, 2, 2.5, 2.5, 2, 1.5, 2, 2, 2, 2))
                
            Case "fema disaster impact"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("cyclone", "disaster", "disaster-related impacts", "earthquake", "environmental hazards", "fema", "fema flood zones", "fire", "fires", "flood", "flooding", "floodplains", "helene", "hurricane", "milton", "storm", "tropical storm", "water-related risks"), _
                Array(2, 2.5, 2, 2, 2, 2.5, 2.5, 2, 2, 2.5, 2.5, 2, 2.5, 2.5, 2.5, 2, 2, 2))
                
            Case "fha/usda requirements"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("4000.1", "case number", "fha", "hud", "mpr", "mps", "usda", "well and septic distance", "hud/fha"), _
                Array(2, 2, 2.5, 2.5, 2, 2, 2.5, 2, 2.5))
                
            Case "handrail"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("handrail", "hand rail", "handrails", "hand rails"), _
                Array(2.5, 2.5, 2.5, 2.5))
                
            Case "income approach"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("1007", "capitalization rates", "income", "landlord", "lease", "operating expenses", "rent", "rental income analysis", "tenant"), _
                Array(2.5, 2, 2.5, 2, 2, 2, 2.5, 2, 2))
                
            Case "inspection requirements"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("builder", "carbon monoxide", "co", "deficiencies", "detectors", "disaster", "double strap", "inspection", "inspection issues", "physical deficiencies", "property evaluation", "smoke", "smoke detectors", "water heater", "co detect"), _
                Array(1.5, 2, 2, 2, 2, 1.5, 2, 2.5, 2, 2, 2, 2, 2, 2, 2))
                
            Case "legal and ownership details"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("apn", "county", "deed", "deed restriction", "descriptions", "legal", "legal definitions", "legal ownership", "occupancy", "oopr", "oor", "owner", "owner of record", "ownership", "pr", "public record", "section", "seller", "tenant", "title", "titled"), _
                Array(2, 1, 2, 2, 1.5, 2, 2, 2, 1.5, 2, 2, 1.5, 2, 2, 2, 2, 1.5, 1.5, 2, 2, 2))
                
            Case "map/exhibits"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("aerial", "comp map", "comparable map", "exhibit", "flood map", "location map", "maps", "plat", "survey"), _
                Array(2, 2, 2, 2, 2, 2, 2, 2, 2))
                
            Case "market trends"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("change in value", "decrease", "decreasing", "home price index", "hpi", "increase", "increasing", "market trend", "market trends", "price changes", "stable market adjustment", "time adjustment", "marketability"), _
                Array(2, 2, 2, 2, 2, 2, 2, 2.5, 2.5, 2, 2, 2, 2))
                
            Case "neighborhood description"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("area details", "boundaries", "built-up", "busy", "community", "demand", "demand/supply", "external factors", "location", "neighborhood", "neighborhood name", "over improved", "private road", "rural", "suburban", "supply", "surroundings", "under improved"), _
                Array(1.5, 1.5, 1.5, 1.5, 2, 1.5, 1.5, 1.5, 2, 2.5, 2, 1.5, 1.5, 1.5, 1.5, 1.5, 2, 1.5))
                
            Case "new construction"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("budget", "builder", "builder name", "builders name", "builder's name", "new construction", "plans", "plans and specs", "specs"), _
                Array(1.5, 2, 2, 2, 2, 2.5, 2, 2, 2))
                
            Case "photos and documentation"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("album", "archive", "blur the people", "blur the person", "blur the photo", "camera", "capture", "catalog", "collection", "compilation", "depiction", "documentation", "dossier", "evidence", "file", "folder", "frame", "gallery", "illustration", "image", "imagery", "images", "index", "inventory", "log", "photo", "photograph", "photos", "pic", "pictorial", "picture", "pictures", "plat map", "portfolio", "portrait", "proof", "record", "register", "repository", "representation", "scene", "shot", "snapshot", "still", "view", "visual", "visual aid", "photographs"), _
                Array(1.5, 1.5, 2, 2, 2, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 2.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 2.5, 2, 2.5, 1.5, 1.5, 1.5, 2.5, 2.5, 2.5, 1.5, 1.5, 2, 2, 2, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 2, 2, 2.5))
                
            Case "property description"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("boundary disputes", "condition", "easement", "easements", "encroachment", "encroachments", "finished", "finishes", "location", "lot size", "parcel", "quality", "rating", "right-of-way issues", "site area", "site size", "structure", "view", "year built"), _
                Array(1.5, 2, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 2, 1.5, 1.5, 1.5, 1.5, 2, 1.5, 2))
                
            Case "property taxes and assessments"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("assessed values", "assessor", "estimated taxes", "property tax rates", "real estate taxes", "special assessment", "tax assessment", "tax assessments", "taxes"), _
                Array(2, 2, 2, 2.5, 2, 2, 2.5, 2.5, 2))
                
            Case "pud/hoa information"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("annual", "hoa", "hoa fee", "hoa fees", "hoa-related information", "homeowners association", "monthly", "planned unit", "pud", "pud details"), _
                Array(1.5, 2.5, 2, 2, 2, 2, 1.5, 2, 2.5, 2))
                
            Case "purchase price"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("purchase price", "sale price", "sales price"), _
                Array(2.5, 2.5, 2.5))
                
            Case "re-built?"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("can subject be re-built to its current footprint if destroyed"), Array(2.5))
                
            Case "repairs and maintenance"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("asbestos", "broken", "contamination", "cost to cure", "ctc", "damage", "damaged", "deferred", "environmental", "hazard", "hazardous materials", "improvements", "lead", "lead paint", "maintain", "maintained", "maintenance", "moisture", "mold", "repair", "repaired", "repairs", "sor", "sow", "toxic", "water", "workmanlike", "workman-like"), _
                Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2.5, 2, 2, 2.5, 2, 2.5, 2, 2, 2, 2, 2, 2))
                
            Case "roof related"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("roof"), Array(2.5))
                
            Case "sales comparison approach/ adjustments"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("above grade", "access", "adjusted", "adjustment", "adjustment grids", "below grade", "bracket", "bracketed", "bracketing", "comp", "comparable", "comparable properties", "comparable sale", "comparable sales", "distance", "gross adjustment", "gross net line", "line item excessive", "market analysis", "mile", "narrative", "net adjustment", "parameter", "parameters", "proximate", "proximity", "sales", "search", "un bracketed", "unadjusted", "unsupported", "comps", "comparables", "adjustments"), _
                Array(2, 1.5, 2, 2.5, 2, 2, 2, 2, 2, 2.5, 2.5, 2, 2.5, 2.5, 1.5, 2, 2, 2, 2, 1.5, 1.5, 2, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 2, 2, 2, 2.5, 2.5, 2.5))
                
            Case "signature"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("signature"), Array(2))
                
            Case "sketch and floor plan issues"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("ansi", "architectural drawing", "arrangement", "blueprint", "configuration", "depiction", "design", "diagram", "floor plan", "floorplan", "gla", "gross living area", "label", "layout", "layout problem", "layout problems", "outline", "plan", "property dimensions", "rendering", "schematic", "scheme", "sketch"), _
                Array(2, 2, 1.5, 2, 1.5, 1.5, 1.5, 1.5, 2.5, 2.5, 2, 2, 1.5, 2, 2, 2, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 2.5))
                
            Case "subject to"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("subject to"), Array(2.5))
                
            Case "tax doc"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("tax"), Array(2))
                
            Case "utility information"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("connection", "covenant", "electricity", "gas services", "maintenance agreement", "oil", "private", "private road", "private street", "private well", "propane", "public utilities", "road", "septic", "sewer", "solar", "solar panels", "street", "utility", "water", "water supply", "wells", "utilities"), _
                Array(2, 1.5, 2, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 2, 1.5, 2, 1.5, 2, 2, 1.5, 1.5, 1.5, 2, 2, 2, 1.5, 2))
                
            Case "value conclusion"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("reconciliation", "value", "value conclusion"), _
                Array(2, 2.5, 2.5))
                
            Case "zoning and highest and best use (hbu)"
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("55+", "accessory unit", "adu", "age restricted", "coach house", "commercial", "conform", "conforming", "conforms", "farm", "guest house", "hbu", "highest and best use", "illegal", "in law", "in-law", "land use", "land use restrictions", "legality", "nonconfirming", "non-conforming", "permit", "permits", "permitted", "permitted uses", "permitting", "related living", "senior", "survey", "zone", "zoned", "zoning", "zoning classifications", "zoning land use"), _
                Array(2, 2, 2, 2, 1.5, 1.5, 2, 2, 2, 1.5, 1.5, 2.5, 2.5, 2, 1.5, 1.5, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1.5, 1.5, 2, 1.5, 1.5, 2.5, 2, 2))
                
            Case Else
                ' For any remaining categories, create minimal keyword/weight arrays
                Call AddUniqueKeywords(tempKeywords, tempWeights, _
                Array("general"), Array(1))
        End Select
        
        ' Convert collections to arrays
        If tempKeywords.Count > 0 Then
            ReDim cats(i).keywords(1 To tempKeywords.Count)
            ReDim cats(i).weights(1 To tempKeywords.Count)
            
            For j = 1 To tempKeywords.Count
                cats(i).keywords(j) = tempKeywords(j)
                cats(i).weights(j) = tempWeights(j)
            Next j
        Else
            ' Empty arrays for categories with no keywords
            ReDim cats(i).keywords(1 To 1)
            ReDim cats(i).weights(1 To 1)
            cats(i).keywords(1) = ""
            cats(i).weights(1) = 0
        End If
    Next i
    
    ' Process each cell in column G
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    totalCount = 0
    categorizedCount = 0
    
    For Each cell In ws.Range("G2:G" & lastRow)
        If Not IsEmpty(cell) Then
            totalCount = totalCount + 1
            
            If VarType(cell.value) = vbString And Not IsError(cell.value) Then
                Dim cellText As String
                cellText = LCase(Trim(cell.value))
                
                ' Find best weighted category
                Dim bestCategory As String
                Dim bestScore As Single
                bestCategory = FindBestWeightedCategory(cellText, cats, bestScore)
                
                If bestCategory <> "" Then
                    cell.Offset(0, 5).value = bestCategory ' Output to column L (G + 5)
                    categorizedCount = categorizedCount + 1
                Else
                    cell.Offset(0, 5).value = "No Primary noted"
                End If
            Else
                cell.Offset(0, 5).value = "No Primary noted"
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Show completion message
    Dim uncategorizedCount As Long
    uncategorizedCount = totalCount - categorizedCount
    
    If totalCount > 0 And (uncategorizedCount / totalCount) > 0.1 Then
        MsgBox "Warning: More than 10% uncategorized. Consider refining keywords.", vbExclamation
    End If
    
    MsgBox "Weighted categorization complete!" & vbCrLf & _
           "Categorized: " & categorizedCount & " of " & totalCount & vbCrLf & _
           "Uncategorized: " & uncategorizedCount, vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error " & Err.Number & ": " & Err.Description & " at line " & Erl, vbCritical
End Sub

' Helper function to find best weighted category
Private Function FindBestWeightedCategory(cellText As String, cats() As Category, ByRef bestScore As Single) As String
    Dim i As Integer, j As Integer
    Dim currentScore As Single
    Dim tempBestCategory As String
    
    bestScore = 0
    tempBestCategory = ""
    
    For i = 1 To UBound(cats)
        currentScore = 0
        For j = 1 To UBound(cats(i).keywords)
            If cats(i).keywords(j) <> "" Then
                If InStr(cellText, LCase(Trim(cats(i).keywords(j)))) > 0 Then
                    currentScore = currentScore + cats(i).weights(j)
                End If
            End If
        Next j
        
        If currentScore > bestScore Then
            bestScore = currentScore
            tempBestCategory = cats(i).Name
        End If
    Next i
    
    FindBestWeightedCategory = tempBestCategory
End Function

' Helper function to add unique keywords and weights, preventing duplicates
Private Sub AddUniqueKeywords(keywordCollection As Collection, weightCollection As Collection, keywords As Variant, weights As Variant)
    Dim i As Integer
    Dim keyword As String
    Dim isDuplicate As Boolean
    
    For i = LBound(keywords) To UBound(keywords)
        keyword = LCase(Trim(CStr(keywords(i))))
        isDuplicate = False
        
        ' Check for duplicates in collection
        Dim j As Integer
        For j = 1 To keywordCollection.Count
            If LCase(Trim(CStr(keywordCollection(j)))) = keyword Then
                isDuplicate = True
                Exit For
            End If
        Next j
        
        ' Only add if not duplicate and not empty
        If Not isDuplicate And keyword <> "" Then
            keywordCollection.Add keywords(i)
            If i <= UBound(weights) Then
                weightCollection.Add weights(i)
            Else
                weightCollection.Add 1# ' Default weight
            End If
        End If
    Next i
End Sub

' Get unique categories from your CSV data
Private Function GetUniqueCategories() As Collection
    Dim uniqueCats As New Collection
    Dim categories As Variant
    
    ' All unique categories from your CSV file
    categories = Array( _
        "Address", "Ammenities", "Appraiser Details and Credentials", _
        "Assignment Type", "Attic related", "Borrower and Lender Information", _
        "closing cost", "Compliance with Standards", "Condo Project Information", _
        "Contract Section", "Cost Approach", "Discrepancies and Inconsistencies", _
        "FEMA Disaster Impact", "FHA/USDA Requirements", "handrail", _
        "Income Approach", "Inspection Requirements", "Legal and Ownership Details", _
        "Map/Exhibits", "Market Trends", "Neighborhood Description", _
        "New Construction", "Photos and Documentation", "Property Description", _
        "Property Taxes and Assessments", "PUD/HOA Information", "Purchase Price", _
        "re-built?", "Repairs and Maintenance", "Roof Related", _
        "Sales Comparison Approach/ Adjustments", "Signature", _
        "Sketch and Floor Plan Issues", "Subject To", "Tax Doc", _
        "Utility Information", "Value Conclusion", "Zoning and Highest and Best Use (HBU)" _
    )
    
    Dim i As Integer
    For i = LBound(categories) To UBound(categories)
        On Error Resume Next
        uniqueCats.Add categories(i), CStr(categories(i))
        On Error GoTo 0
    Next i
    
    Set GetUniqueCategories = uniqueCats
End Function
