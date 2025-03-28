# Introduction
📊This is a data analysis project that consists in collecting data regarding real estate internet postings in Warsaw and surrounding counties. I analyzed over 22000 offers and presented the results visually in Power BI.

The final product is a tool that can be used by anyone looking for a house or an appartment. It gathers all data from 4 different websites in one place and enables searching for offers based on location and comparing different offers.

😎Here is a sneak-peak of the final dashboard:

![Dashboard](/Assets/Photos/Dashboard.png)

Below you can find a short version of the documentation. To see the full documentation go to [Project Documentation](/Project%20Documentation%20(README)%20Real%20Estate%20Market%20Analysis.pdf).

# Tools I Used
In this project I used a combination of Power Query and VBA to scrape the data from the internet and to automate this process. Then, I cleaned the data using Power Query and loaded it to Power BI. The results are additionally summarized in form of a Power Point presentation.

🛜Here is one part of the web scraping M Code:

```

let
    // PART 1 - WEB SCRAPING

    // Step 1: Define the base URL for the website
    Location = Excel.CurrentWorkbook(){[Name="Otodom_Setup"]}[Content]{0}[Column1],
    BaseUrl = Text.From("https://www.otodom.pl/pl/wyniki/wynajem/mieszkanie/" & Location & "?limit=72&by=DEFAULT&direction=DESC&viewType=listing&page="),
    
    // Step 2: Extract the total number of pages dynamically
    SourcePage = Web.BrowserContents(BaseUrl & "1"),  // Load the first page to extract pagination info
    TotalResultsText = Table.SelectRows(
    Html.Table(
        SourcePage, {{"Pagination", ".css-15svspy"}}),
    each not Text.StartsWith([Pagination],"Jak")),
    TotalPages = Number.RoundUp(
    Number.FromText(
        Text.AfterDelimiter(TotalResultsText{0}[Pagination], " z "))/72),  // Convert the text to a number
    
    // Step 3: Generate a list of URLs for all pages
    PageNumbers = List.Numbers(1, if TotalPages > 15 then 15 else TotalPages),  // Generate a list [1, 2, ..., TotalPages]
    Urls = List.Transform(PageNumbers, each BaseUrl & Text.From(_)),  // Create URLs by appending page numbers to the base URL
    
    // Step 4: Define a function to fetch data from a single page
    FetchPage = (url as text) =>
        try
            let
                Source = Web.BrowserContents(url),
                Data = Html.Table(Source, {{"Cena", ".css-2bt9f1"}, {"Opis", ".css-u3orbr"}, {"Adres", ".css-42r2ms"}, {"Pokoje", "DD:nth-child(2)"}, {"Metraż", "DD:nth-child(4)"}, {"Piętro", "DD:nth-child(6)"}, {"Czynsz", ".css-13du2ho"}, {"Wynajmujący", ".css-1sylyl4"}, {"Biuro nieruchomości", ".css-196u6lt"}}, [RowSelector=".css-2bt9f1"])
            in
                Data
        otherwise
            null,  // Return null if the page fails to load
    
    // Step 5: Apply the function to all URLs and fetch data
    AllData = List.Transform(Urls, each FetchPage(_)),  // Fetch data from each page
    
    // Step 6: Combine data from all pages into a single table
    ValidData = List.RemoveNulls(AllData),  // Remove null results (in case some pages failed)
    CombinedData = Table.Combine(ValidData),  // Combine all page data into one table

    // PART 2 - DATA CLEANING

    // Define the list of replacement rules
    Replacements = {
        { " m²", "", {"Metraż"} },
        { " piętro", "", {"Piętro"} },
        { "parter", "1", {"Piętro"} },
        { "Biuro nieruchomości", "true", {"Biuro nieruchomości"} },
        { ".", ",", {"Metraż"} },
        { "zł/miesiąc", "", {"Cena"} }
    },

    // Apply replacements iteratively using List.Accumulate
    ReplacedValues = List.Accumulate(
        Replacements, 
        CombinedData, 
        (table, replacement) => 
            Table.ReplaceValue(
                table, 
                replacement{0}, 
                replacement{1}, 
                Replacer.ReplaceText, 
                replacement{2})),
    ExtractedPokoje = Table.TransformColumns(ReplacedValues, {{"Pokoje", each Text.Start(_, 1), type text}}),
    ExtractedCena = Table.TransformColumns(ExtractedPokoje, {{"Cena", each Text.BeforeDelimiter(_, "zł"), type text}}),
    ExtractedCzynsz = Table.TransformColumns(
    ExtractedCena,
    {{"Czynsz", each Text.AfterDelimiter(
        Text.BeforeDelimiter(_, "zł/"),
        "czynsz:")}}),
    AddedDate = Table.AddColumn(ExtractedCzynsz, "Data pobrania", each Date.From(DateTime.LocalNow()), type date),
    AddedLocation = Table.AddColumn(AddedDate, "Lokalizacja-link", each Location, type text),

    // PART 3 - DATA LOAD

    HistoricalData = Table.PromoteHeaders(
        Excel.Workbook(
            File.Contents("C:\Users\Marcin\Desktop\Varia\Praca\Portfolio\Mieszkania\Mieszkania - baza danych.xlsm"), 
            null, 
            true
        ){[Item="M1_Otodom",Kind="Sheet"]}[Data], 
        [PromoteAllScalars=true]),
    Append = Table.Combine({AddedLocation, HistoricalData}),
    RemovedBlankRows = Table.SelectRows(Append, each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),

    // Error handling in case of wrong data types 
    TypePrep = Table.TransformColumns(
        RemovedBlankRows,
        {
            {"Cena", each try Number.From(_) otherwise null},
            {"Pokoje", each try Number.From(_) otherwise null},
            {"Metraż", each try Number.From(_) otherwise null},
            {"Piętro", each try Number.From(_) otherwise null},
            {"Czynsz", each try Number.From(_) otherwise null}}),
    ChangedType = Table.TransformColumnTypes(TypePrep,{{"Cena", type number}, {"Opis", type text}, {"Adres", type text}, {"Pokoje", Int64.Type}, {"Metraż", type number}, {"Piętro", Int64.Type}, {"Czynsz", type number}, {"Wynajmujący", type text}, {"Biuro nieruchomości", type logical}, {"Data pobrania", type date}, {"Lokalizacja-link", type text}}),
    ReplacedErrors = Table.ReplaceErrorValues(ChangedType, {{"Cena", null}, {"Opis", null}, {"Adres", null}, {"Pokoje", null}, {"Metraż", null}, {"Piętro", null}, {"Czynsz", null}, {"Wynajmujący", null}, {"Biuro nieruchomości", null}, {"Data pobrania", null}}),
    RemovedDuplicates = Table.Distinct(ReplacedErrors, {"Cena", "Opis", "Adres", "Pokoje", "Metraż", "Piętro", "Czynsz", "Wynajmujący", "Biuro nieruchomości"})
in
    RemovedDuplicates

```

💻And here comes the VBA automation macro:

```VB

Sub RefreshQueries()

Dim i As Integer
Dim LR As Integer
Dim Location As String
Dim FstTime, SndTime, TrdTime

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.EnableAnimations = False

LR = Setup_M4.Range("A1", Setup_M4.Range("A1").End(xlDown)).Rows.Count

Setup_M4.Range("A2:A" & LR).Copy Setup_M4.Range("A" & LR + 2)

'Loop
On Error Resume Next
For i = 2 To LR
    
    FstTime = Timer
    
    'Refresh the query
    ThisWorkbook.Queries.FastCombine = True
    ThisWorkbook.Connections("Query - Rentola").Refresh
    
    SndTime = Timer
    Location = Setup_M4.Range("A2")
    
    'Delete the last used parameter in Setup
    Setup_M4.Range("A2").ListObject.ListRows(1).Delete
    
    'Save the workbook to be able to use the next query correctly
    ThisWorkbook.Save
    
    'Wait 1 minute to avoid sending too many requests to the website and getting blocked
    Application.Wait (Now + TimeValue("0:01:00"))
    TrdTime = Timer
    
    'Debug
    Debug.Print i & " / Refresh Time: " & SndTime - FstTime & " / Wait Time: " & TrdTime - SndTime & " / " & Location

Next i

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.EnableAnimations = True

End Sub

```

# The Analysis

Although the main goal was to give the end user a tool that would enable them to explore the data on their own, I still believe it is a good idea at least to summarize the results:

### 1. The most offers of apartments for rent came from the Warsaw districts Mokotów and Śródmieście. The most offers of houses for sale were to find in powiat piaseczyński.

### 2. In terms of apartments, the highest rental prices per square meter are in the districts of Wola and Śródmieście.

### 3. The prices per square meter of houses in the suburbs of Warsaw are very close or even lower than in some towns in the surrounding counties.

![Insights](/Assets/Photos/Slide3.JPG)

For more details I recommend you checking out the [presentation](/Analiza%20ogłoszeń%20rynku%20nieruchomości.pdf).📈📊

# What I Learned

Working on this project gave me an experience in web scraping which was the most challenging part of it. Among many other things, I learned how to imitate a regular user by forcing my code to wait some time to avoid sending too many requests. This, of course, raised my competences in the Power Query M language and VBA even higher.💪

Additionally, while building the dashboard I refreshed some useful concepts of Power BI.

🔗To see the full list of skills engaged in this project I recommend you to visit my [LinkedIn profile](https://www.linkedin.com/in/marcin-czerkas-95150727a/).

# Conclusion

Finally, what is the added value of my project?

I developed a tool that might help those who are looking for an apartment to rent or a house to buy. Of course, it is not an exhaustive market analysis. However, the project focuses on the data that such person would like to check anyway – but much slower by checking all websites manually. Thanks to the tool I created all offers are gathered in one place and visualized in a clear way.

🎉It has been my second data project. The first one was a Power BI dashboard build on top of a SQL database created by me to follow the results of a tabletop game. If you are interested, [check it out](https://github.com/MarcinCzerkas/Project-Middle-earth-SBG)!
