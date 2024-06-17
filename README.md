## Loading Senior Theses into Alma Digital using Google Forms, Apps Script, and Alma

[Public Theses Site](https://primo.lclark.edu/discovery/collectionDiscovery?vid=01ALLIANCE_LCC:LCC&inst=01ALLIANCE_LCC&collectionId=81244751090001844&sortby=title&sortItemsBy=date_d)

### Macro Steps

1. **Collect files and metadata from students via Google Forms**
2. **Run Google Apps Script to create MarcXML, rename files, and download zip to your computer**
3. **In Alma, create an Ingest, and run the Import Profile to create records**
4. **Spot check to make sure it worked**

**Note:** One future direction could be to use Alma APIs to submit the ingest directly from the Google Spreadsheet, with Apps Script. Who knows?

### Step by Step

#### Google Forms/Spreadsheets Steps
1. **Set up Google Form for a given semester.** [This](https://forms.gle/RXB2As71QmNCAQWAA) is a good model to copy from. Make sure responses are saved to a spreadsheet, and “total uploaded files” is set to 100Gb. This can’t live in the Watzek shared drive, since it doesn’t allow file uploads.
2. **Share the Form with the faculty listserv (cas-faculty@lclark.edu), maybe a month before semester’s end.** When sharing, offer to share the spreadsheet with thesis advisors, so they can keep track of what their students submitted.
3. **When time to process (typically a couple weeks or so after semester’s end), add the [Google Apps Script code](googleAppsScript.js) from our GitHub repository (it’s private, requires GitHub account and Watzek GitHub membership) to the spreadsheet:**
    - Go to `Extensions -> Apps Scripts`. Paste the code from the link above, and save.
    - In the Apps Scripts menu on the left side, click `Triggers`.
    - Click `Add Trigger` (lower right).
    - **Settings:**
      - Function to Run: `onOpen`
      - Deployment: `Head`
      - Event Source: `from spreadsheet`
      - Select event type: `On Open`
    - Save (you may have to allow your popup blocker, and provide google authorization).
    - Refresh the spreadsheet, and you should see the menu option `Thesis Functions`.
4. **Click `Thesis Functions -> Prepare metadata for XML`.** This adds columns at the end with values that will populate certain Marc record fields, and renames Google Drive files using our naming convention.
5. **Click `Thesis Functions -> Generate Marc XML`.** This creates a Marc XML file from the spreadsheet data, and saves it in the same Google Drive folder as the thesis files.
6. **Click `Thesis Functions -> Download Zip for Alma Ingest`.** This will create a zipped file of the thesis files and Marc XML, and provide a link to it from a popup box. Click the link, and then click the icon to download the zip file to your computer. (Note: this zip file is created in your Google Drive account’s root level. You may want to delete it later, to save space).
7. **On your computer, unzip the downloaded zip file.**

#### Alma Steps
1. **In Alma, go to `Resources -> Advanced Tools -> Digital Uploader`.**
2. Click `Add New Ingest`.
3. Give the Ingest a name (e.g., “spring2024”), and click `Add files`, which will prompt you to upload files from your computer. Find the unzipped folder, and select all files inside (pdfs plus the one xml). Once they all appear in the table, click `Upload All`, and then `OK` when finished.
4. **Back in the Digital Uploader menu, select the checkbox next to the Ingest you created, and click `Submit Selected`.** This will enable the `Run MD Import` button. Make sure the `Lewis & Clark Senior Thesis Collection (Student Theses Marc21XML profile)` Import Profile is selected, and click the `Run MD Import` button. After a few seconds, you should see a feedback window on the right indicating that the Import has been submitted.
5. **You can monitor the input at `Resources -> Input -> Monitor and View Inputs`.** If it failed, try to dig through the reports/events to find some evidence why - chances are it had to do with the XML file. Sometimes the feedback is lacking, so it might be worth trying to check with an XML validator. If it worked, you might click the three dots on the row of your import, and click `Imported Records`. Ideally the list should match what’s in the Google Spreadsheet.
6. **Spot check your work.** The records should be in Alma right away, but it takes a few minutes or more to publish to the Senior Thesis collection. Using the spreadsheet as a guide, take a look at some of the new records. How does the metadata look? Are the Access Rights Policies working?

## Additional Notes

**Alma Digital Training Videos:** [Alma Digital Training Videos](https://exlibris.libguides.com/alma/almadigital)
There are many, but it provides good background on all the pieces of Alma Digital, loading content, Import Profiles, etc.

### Collection Structure in Alma Digital

In Alma, go to `Resources -> Manage Inventory -> Manage Collections`.
There is a top-level collection titled Lewis & Clark Senior Thesis Collection. This collection contains many sub-collections - one for each department (e.g., Art History, Asian Studies, Biology, etc.).

### Access Rights Policies

Alma Access Rights Policies determine which users in Primo can access content. These are defined in `Configuration -> Fulfillment -> Copyright Management -> Access Rights`. Currently we use two policies:
- “Open to the world” - does not require login to access Alma Digital File
- “CAS access” - includes undergrads and CAS Faculty Staff - all others denied.

To associate an Access Rights Policy to an Alma Digital record, the name of the Access Rights Policy must be included in the MarcXML 540a field:

```xml
<marc:datafield tag="540" ind1=" " ind2=" ">
    <marc:subfield code="a">Open to the world</marc:subfield>
</marc:datafield>
```

or 

```xml
<marc:datafield tag="540" ind1=" " ind2=" ">
    <marc:subfield code="a">CAS access</marc:subfield>
</marc:datafield>
```


The Google Form provides a choice for users to submit their theses: “LC - only” or “public”. The Google Spreadsheet apps script has a function to rename those values to map to the Access Rights Policy names.
 

### Deleting Records
There may be times you want to delete Alma Digital records from Alma. There may be a better way to do this, but here’s one: In Alma, go to Resources->Imports->Manage Imports. Select the batch import, and click “View Records”. For each record, you can then delete the digital representation. When doing so, you’ll be prompted to delete the bib record as well, which you should.
