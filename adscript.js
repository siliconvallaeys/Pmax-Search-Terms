function main() {
    /******************************************
     * PMax Search Terms Report
     * @version: 1.0
     * @authors: Frederick Vallaeys (Optmyzr)
     * -------------------------------
     * Install this script in your Google Ads account (not an MCC account)
     * to generate a spreadsheet containing the search terms in your Performance Max campaigns.
     * The spreadsheet also includes data about category labels (groupings of search terms).
     * Metrics include conversion value, conversions, clicks, and impressions
     * --------------------------------
     * v2 (Mar 14, 2024): adds a new column that says if the search terms exists as a keyword in a search campaign. adds better error handling.
     * --------------------------------
     * For more PPC tools and scripts, visit www.optmyzr.com.
     ******************************************/

    var minImp = 10; // Limit the output to only items with at least this many impressions
    var spreadsheetUrl = ""; // leave blank to generate a new spreadsheet or add your own URL to overwrite the data in an existing spreadsheet
    var reportLastNDays = 30; // The number of days to include in the report
    var EMAILADDRESS = ""; // enter your email address

    // Don't edit below this line unless you know how to write scripts
    //-----------------------------------------------------------------
    let pmaxOnly = 1;

    let allCategoryLabels = [
        [
            "Campaign Name",
            "Category Label",
            "Conv Val",
            "Conv",
            "Clicks",
            "Imp",
        ],
    ];
    let allSearchTerms = [
        [
            "Campaign Name",
            "Category Label",
            "Subcat",
            "Search Term",
            "Conv Val",
            "Conv",
            "Clicks",
            "Imp",
            "Exists as Keyword?",
        ],
    ];

    let keywordsList = getAllSearchCampaignKeywords();

    var dateRange = getDateRange(reportLastNDays);

    try {
        let baseQuery = `
    SELECT 
    campaign.id, 
    campaign.name, 
    metrics.clicks, 
    metrics.impressions, 
    metrics.conversions, 
    metrics.conversions_value
    FROM campaign
    WHERE campaign.status != 'REMOVED'
    AND metrics.impressions > 0
    AND segments.date BETWEEN ${dateRange}
    AND metrics.impressions >= ${minImp}
    `;

        // Modify the SQL query if pmaxOnly is true
        if (pmaxOnly) {
            baseQuery +=
                " AND campaign.advertising_channel_type = 'PERFORMANCE_MAX' ";
        }

        baseQuery += "ORDER BY metrics.conversions DESC";

        let campaignIdsQuery = AdsApp.report(baseQuery);

        let rows = campaignIdsQuery.rows();

        while (rows.hasNext()) {
            let campaignRow = rows.next();
            let campaignId = campaignRow["campaign.id"];
            Logger.log(
                campaignRow["campaign.id"] + " " + campaignRow["campaign.name"]
            );

            // Search Labels Report
            let categoryLabelsQuery = `
      SELECT 
                campaign_search_term_insight.category_label, 
                campaign_search_term_insight.id,
                metrics.clicks, 
                metrics.impressions, 
                metrics.conversions,
                metrics.conversions_value
              FROM 
                campaign_search_term_insight 
              WHERE 
                segments.date BETWEEN ${dateRange}
                AND campaign_search_term_insight.campaign_id = '${campaignId}'
                AND metrics.impressions >= ${minImp}
              ORDER BY 
                metrics.conversions DESC
              `;
            let categoryLabelsQueryResult = AdsApp.report(categoryLabelsQuery);
            let categoryLabelsResults = categoryLabelsQueryResult.rows();
            while (categoryLabelsResults.hasNext()) {
                let categoryLabelsRow = categoryLabelsResults.next();
                let categoryLabelId =
                    categoryLabelsRow["campaign_search_term_insight.id"];
                //Logger.log(categoryLabelId + ". " + categoryLabelsRow['campaign_search_term_insight.category_label'] + " " + categoryLabelsRow['metrics.impressions']);
                allCategoryLabels.push([
                    campaignRow["campaign.name"],
                    categoryLabelsRow[
                        "campaign_search_term_insight.category_label"
                    ],
                    categoryLabelsRow["metrics.conversions_value"].toFixed(2),
                    categoryLabelsRow["metrics.conversions"].toFixed(1),
                    categoryLabelsRow["metrics.clicks"],
                    categoryLabelsRow["metrics.impressions"],
                ]);
                // Search Terms Report
                let searchTermsQuery = `
        SELECT 
                  metrics.clicks, 
                  metrics.impressions, 
                  metrics.conversions,
                  metrics.conversions_value,
                  segments.search_term,
                  segments.search_subcategory
                FROM 
                  campaign_search_term_insight 
                WHERE 
                  segments.date BETWEEN ${dateRange}
                  AND campaign_search_term_insight.campaign_id = '${campaignId}'
                  AND campaign_search_term_insight.id = '${categoryLabelId}'
                `;
                let searchTermsQueryResult = AdsApp.report(searchTermsQuery);
                let searchTermsResults = searchTermsQueryResult.rows();
                // Inside the loop where search terms are being processed
                while (searchTermsResults.hasNext()) {
                    let searchTermsRow = searchTermsResults.next();
                    if (searchTermsRow["metrics.impressions"] >= minImp) {
                        let searchTermText = searchTermsRow[
                            "segments.search_term"
                        ].toLowerCase(); // Normalize search term text for comparison
                        let existsAsKeyword = keywordsList[searchTermText]
                            ? "Yes"
                            : "No";

                        allSearchTerms.push([
                            campaignRow["campaign.name"],
                            categoryLabelsRow[
                                "campaign_search_term_insight.category_label"
                            ],
                            searchTermsRow["segments.search_subcategory"],
                            searchTermsRow["segments.search_term"],
                            searchTermsRow["metrics.conversions_value"].toFixed(
                                2
                            ),
                            searchTermsRow["metrics.conversions"].toFixed(1),
                            searchTermsRow["metrics.clicks"],
                            searchTermsRow["metrics.impressions"],
                            existsAsKeyword, // New column indicating if the search term exists as a keyword
                        ]);
                    }
                }
            }
        }

        if (!spreadsheetUrl) {
            var ss = SpreadsheetApp.create("PMax Search Terms", 10000, 20);
            var spreadsheetUrl = ss.getUrl();
        } else {
            var ss = SpreadsheetApp.openByUrl(spreadsheetUrl);
        }

        let categoriesSheet = ss.getSheetByName("categories")
            ? ss.getSheetByName("categories").clear()
            : ss.insertSheet("categories");
        if (allCategoryLabels.length > 1) {
            // Check if there's more than just the header row
            categoriesSheet
                .getRange(
                    1,
                    1,
                    allCategoryLabels.length,
                    allCategoryLabels[0].length
                )
                .setValues(allCategoryLabels);
        }

        let termsSheet = ss.getSheetByName("terms")
            ? ss.getSheetByName("terms").clear()
            : ss.insertSheet("terms");
        if (allSearchTerms.length > 1) {
            // Check if there's more than just the header row
            termsSheet
                .getRange(1, 1, allSearchTerms.length, allSearchTerms[0].length)
                .setValues(allSearchTerms);
        }

        var subject = "PMax Search Terms Report Ready";
        var body =
            "The PMax Search Terms Report has been generated and is available at: " +
            spreadsheetUrl +
            "\n\nReport covers the last " +
            reportLastNDays +
            " days." +
            "\n\nThis is an automated email sent by Google Ads Script.";
    } catch (e) {
        Logger.log("Error: " + e.message);
        var subject = "PMax Search Terms Report Failed";
        var body =
            "The PMax Search Terms Report encountered an error: " +
            e +
            "\n\nThis is an automated email sent by Google Ads Script.";
    }
    Logger.log("spreadsheet: " + spreadsheetUrl);

    // Send the email
    var recipientEmail = EMAILADDRESS;
    MailApp.sendEmail(recipientEmail, subject, body);
}

// function to get date range
function getDateRange(numDays) {
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(endDate.getDate() - numDays);
    const format = (date) =>
        Utilities.formatDate(
            date,
            AdsApp.currentAccount().getTimeZone(),
            "yyyyMMdd"
        );
    return `${format(startDate)} AND ${format(endDate)}`;
}

function getAllSearchCampaignKeywords() {
    let keywordsList = {};
    let keywordIterator = AdsApp.keywords()
        .withCondition("Status = ENABLED")
        .withCondition("CampaignStatus = ENABLED")
        .withCondition("AdGroupStatus = ENABLED")
        .get();

    while (keywordIterator.hasNext()) {
        let keyword = keywordIterator.next();
        let keywordText = keyword.getText().toLowerCase(); // Normalize keyword text for comparison
        keywordsList[keywordText] = true; // Use a map for efficient lookup
    }

    return keywordsList;
}
