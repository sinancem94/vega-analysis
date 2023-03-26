// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// No Node.js APIs are available in this process unless
// nodeIntegration is set to true in webPreferences.
// Use preload.js to selectively enable features
// needed in the renderer process.

type FieldMap = {
    [key: string]: string[];
};

class AnalysisStage {
    static Import = "Import";
    static SelectSheet = "SelectSheet";
    static SelectFields = "SelectFields";
    static Analyzed = "Analyzed";
    static Exported = "Exported";
}


const versions = window.versions;
const available_formats = [ ".xlsx", ".csv"];
const select_fields: FieldMap = {'part_number':["part number", "pn"], 'part_description':["part description", "part desc", "description", "desc"], 'part_quantity':["part quantity", "part qty", "quantity", "qty"], 
                'unit_prices':["part price", "unit price", "price"], 'unit_currency': ["currency", "curr", "cur"],
                'vendor_code':["vendor code", "vendor"], 'vendor_name':["vendor name", "vendor"], 
                'purchase_order':["purchase order", "po no"], 'order_type':["order type", "po"]};

const filter_fields: FieldMap = {'type_filter': ['order_type'], 'currency_filter': ['unit_currency'], 'vendor_filter': ['vendor_name']};

let selectedSheet: string = null;

///FLOW ////

//const information = document.getElementById('info')
//information.innerText = `This app is using Chrome (v${window.versions.chrome()}), Node.js (v${versions.node()}), and Electron (v${versions.electron()})`;

//const startButton = document.getElementById('start-button');
//startButton.innerHTML = `Import the sheet boss`;
const chooseText = document.getElementById('choose-text');
let upload_name = "upload_link";
chooseText.innerHTML = `Press <span name="${upload_name}" class="link-primary" href="#" >boss</span> to start. Only files with extensions ${generateExtensionString(available_formats)} are allowed.`

let upload_links = document.getElementsByName(upload_name);
for(let i = 0; i < upload_links.length; i++){
    upload_links[i].addEventListener("click", (e:Event) => file_upload());
}

setAnalysisMessage(AnalysisStage.Import);

//// FUNCTIONS


async function file_upload() {
    const sheets = await window.versions.file_upload();
    if (sheets == null)
    {
        return;
    }
    draw_sheet_cols(sheets);
}

function draw_sheet_cols(sheet_names: string[]){
    const sheetTable = document.getElementById('sheet-table');
    
    let sheetDivs = sheetTable.children;
    while (sheetDivs.length > 0) {
        sheetDivs[0].remove();
    }

    var id = 0;
    sheet_names.forEach(name => {
        const sheetDiv = document.createElement("div");
        sheetDiv.classList.add("col", "col-lg-4", "text-center"); // "bg-success", "text-white"
        sheetDiv.textContent = `${name}`;
        sheetDiv.id = `${id}`;
        // add an onclick handler to the sheet div
        sheetDiv.onclick = async () => {

            if(sheetDiv.classList.contains("bg-success"))
            {
                sheetDiv.classList.remove("bg-success", "text-white");
                selectedSheet = null;
            }
            else
            {
                // set the selectedSheet variable to the current sheet number
                selectedSheet = name;
                // update the class of all sheet divs
                const sheetDivs = document.querySelectorAll("#sheet-table > .col");
                sheetDivs.forEach(div => {
                    div.classList.remove("bg-success", "text-white");
                });
                sheetDiv.classList.add("bg-success", "text-white");
            }
            
            await draw_analysis_options(selectedSheet);
            let allSet = await this.draw_analysis_filters(selectedSheet);
            setAnalysisMessage(AnalysisStage.SelectFields, {canAnalyze: allSet});
        };

        // append the sheet div to the sheet table
        sheetTable.appendChild(sheetDiv);
    
        id++;
    });

    setAnalysisMessage(AnalysisStage.SelectSheet);
}

const draw_analysis_options = async (selectedSheet: string) => {
    const columns = await window.versions.get_columns(selectedSheet);
    
    const fieldSelect = document.getElementById('analysis-field-selection');

    let htmlFieldSelect = '';
    htmlFieldSelect += '<br>';


    for (let key in select_fields) {

        htmlFieldSelect += '<div  style="width:90%;">';
        
        const text = key.replace('_', ' ').toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ') + ':';
        htmlFieldSelect += '<label for="' + key + '">' + text + '</label>';
        
        htmlFieldSelect += '<select class="part" onchange="selectFieldChanged()" style="float:right;" aria-label=".form-select-lg example" id="' + key + '" name="' + key + '">';

        let litteralDiff = 20;
        columns.forEach(c => {
            //const hardCodedName = field.replace('_', ' ');
            htmlFieldSelect += '<option value="' + c + '"';
            const index = select_fields[key].findIndex((searchTerm: string) => c.toLocaleLowerCase().includes(searchTerm.toLocaleLowerCase()));
            if (index !== -1) {
                let currLitteralDiff = c.length - select_fields[key][index].length;
                if(currLitteralDiff < litteralDiff){
                    litteralDiff = currLitteralDiff;
                    htmlFieldSelect += ' selected="selected"';
                }
            }
            htmlFieldSelect += '>' + c + '</option>';
        });

        if(4 < litteralDiff){
            htmlFieldSelect += '<option value="" selected="selected" hidden="hidden">Please select</option>';
        }
        
        htmlFieldSelect += '</select>';
        htmlFieldSelect += '</div>'
        htmlFieldSelect += '<br>';
    }
    htmlFieldSelect += '<br>';

    fieldSelect.innerHTML = htmlFieldSelect;
};

const draw_analysis_filters = async () => {
    const analysisFilters = document.getElementById('analysis-field-filters');
    let htmlFilters = '';
    htmlFilters += '<div class="row justify-content-md-center">';

    let allSet = true;

    for (let fieldKey in filter_fields) {
        
        let filterSelect = filter_fields[fieldKey];

        let selected = (document.querySelector('#' + filterSelect + ' option:checked') as HTMLSelectElement).value;
        htmlFilters += '<div class="col col-md-4">';
        const text = fieldKey.replace('_', ' ').toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ') + ':';
        htmlFilters += '<label for="' + fieldKey + '">' + text + '</label>';
        htmlFilters += '<br>';

        let disabled = (selected) ? '' : 'disabled';
        htmlFilters += '<select class="part-filter form-select"  multiple aria-label="multiple select example"' + disabled + ' id="' + fieldKey + '" name="' + fieldKey + '">';

        if (selected) {

            let columnValues = await window.versions.get_column_values_unique(selectedSheet, selected);
            columnValues.forEach(c => {
                //const hardCodedName = field.replace('_', ' ');
                htmlFilters += '<option value="' + c + '"';
                htmlFilters += '>' + c + '</option>';
            });
        }
        else {
            allSet = false;
        }
        
        htmlFilters += '</select>';
        htmlFilters += '</div>';
    }
    htmlFilters += '</div>';
    analysisFilters.innerHTML = htmlFilters;

    return allSet;
}

async function selectFieldChanged(){
    var allSet = await draw_analysis_filters();
    setAnalysisMessage(AnalysisStage.SelectFields, {canAnalyze: allSet});
}

const parts_submit = async () => {

    const parts_map: FieldMap = { };

    for (let key in select_fields) {
        if(!parts_map[key]){
            parts_map[key] = [];
        }

        parts_map[key].push((document.querySelector(`#${key}`) as HTMLInputElement).value);
    }
    parts_map['sheet_name'] = [];
    parts_map['sheet_name'].push(selectedSheet);   

    for (let key in filter_fields) {
        
        let filters: string[] = [];
        let selectedFilters = document.querySelectorAll('#' + key + ' option:checked');

        selectedFilters.forEach(function(selected: any){ 
            filters.push((selected as HTMLSelectElement).value);
        });

        parts_map[key] = filters;
    }

    const analyzeRes = await window.versions.parts_analyse(parts_map);

    if(analyzeRes.success){
        let analysis_text = "Total of " + analyzeRes.uniquePartCount + " unique part analysed.<br>";
        analysis_text += "&ensp;Total of " + analyzeRes.uniqueVendorCount + " unique vendors found.<br>";
        for(const cur in analyzeRes.totalVolume){
            analysis_text += "&ensp;Volume : " + formatMoney(analyzeRes.totalVolume[cur]) + " " + cur + "<br>";
        }
        setAnalysisMessage(AnalysisStage.Analyzed, {message: analysis_text, infoClass:"text-success"});
        
    }
    else{
        setAnalysisMessage(AnalysisStage.SelectFields, {message: "There is an error somewhere.. May be change select fields ?", infoClass:"text-danger"});
    }
}


const analysis_save = async () => {
    const saveResult = await window.versions.save_results();
    if(saveResult.length > 1){
        let analysis_text = "Saved to " + saveResult + " successfully";
        setAnalysisMessage(AnalysisStage.Exported, {message: analysis_text, infoClass:"text-success"});

    }
    else{
        let analysis_text = "Save to " + saveResult + " unsuccessfull :((";
        setAnalysisMessage(AnalysisStage.Exported, {message: analysis_text, infoClass:"text-danger"});

    }
}

//chatGPT
function generateExtensionRegex(extensions: string[]) {
    // Escape all special characters in the extensions
    const escapedExtensions = extensions.map(ext => ext.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    // Join the extensions into a pipe-separated string
    const extensionsString = escapedExtensions.join('|');
    // Create the regular expression
    const regex = new RegExp(`(${extensionsString})$`, 'i');
    // Return the regular expression
    return regex;
}

//ChatGPT
function generateExtensionString(extensions : string[]) {
    // Remove the dot from each extension and join with commas
    return extensions.map(ext => ext.replace(/^\./, '')).join(', ');
}



function setAnalysisMessage(currentStage: string, {message = '', infoClass = '', canAnalyze = false} = {}){

    const analysisMessages = document.getElementById('analysis-messages');

    let firstLine = ':`(';
    let secondClass = '';
    let secondLine = ':`((((';
    let firstButton = '';
    let secondButton = '';
    let thirdLine = '';
    console.log("currentStage : " + currentStage);
    console.log("canAnalyze : " + canAnalyze);

    switch (currentStage) {
        case AnalysisStage.Import:
            firstLine = 'File sececen koyacan zor diil..';
            secondClass = 'text-success';
            secondLine = 'You can do it, I believe in you!!';
            break;
        case AnalysisStage.SelectSheet:
            firstLine = 'Select a sheet (from loaded excel) for analysis';
            secondClass = 'text-success';
            secondLine = 'You can do it!!';
            break;
        case AnalysisStage.SelectFields:
            firstLine = 'Select the fields for analysis. Fields are set from the first row of loaded Excel Sheet.';
            
            if (message != ''){
                secondLine = message;
                secondClass = infoClass;
            }
            else{
                secondLine = '!! Check if the auto placed fields are correct dont be lazy..';
                secondClass = 'text-danger';
            }
            
            if (canAnalyze){
                thirdLine = '<span style="float: right;">All fields are selected. Select filters if you like. <br><span class="small text-warning">Filters not guaranteed. Pay money for guarantee </span></span>';
                firstButton = '<button class="btn btn-success" type="button" onclick="parts_submit()" id="analyse_button">Analyse Parts</button>';
            }
            break;
        case AnalysisStage.Analyzed:
            firstLine = 'You can now press save analysis to export Sheet';
            secondClass = 'text-success';
            secondLine = message;
            firstButton = '<button class="btn btn-success" type="button" onclick="parts_submit()" id="analyse_button">Analyse Parts</button>';
            secondButton = '<button class="btn btn-secondary" onclick="analysis_save()" type="button">Save Analysis</button>';
            break;
        case AnalysisStage.Exported:
            firstLine = 'Wow';
            secondClass = infoClass;
            secondLine = message;
            firstButton = '<button class="btn btn-success" type="button" onclick="parts_submit()" id="analyse_button">Analyse Parts</button>';
            secondButton = '<button class="btn btn-secondary" onclick="analysis_save()" type="button">Save Analysis</button>';
            break;
        default:
            firstLine = 'Dayi bir sikinti var burada';
            secondClass = 'text-danger';
            break;
      }

      analysisMessages.innerHTML = `<div><br><p>&ensp;${firstLine}<br>\
            <span class="${secondClass}">&ensp;${secondLine}</span></p><br>\
            &ensp;${firstButton}&emsp;${secondButton}${thirdLine}</div>`;
}

function formatMoney(amount: number, decimalCount = 2, decimal = ".", thousands = ",") {
    try {
      decimalCount = Math.abs(decimalCount);
      decimalCount = isNaN(decimalCount) ? 2 : decimalCount;
  
      const negativeSign = amount < 0 ? "-" : "";
  
      let i = (+Math.abs(Number(amount) || 0).toFixed(decimalCount)).toString();
      let j = (i.length > 3) ? i.length % 3 : 0;
  
      return negativeSign +
        (j ? i.substr(0, j) + thousands : '') +
        i.substr(j).replace(/(\d{3})(?=\d)/g, "$1" + thousands) +
        (decimalCount ? decimal + Math.abs(amount - parseInt(i)).toFixed(decimalCount).slice(2) : "");
    } catch (e) {
      console.log(e)
    }
}