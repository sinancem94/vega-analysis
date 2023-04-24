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
};

//const versions = window.versions;
const available_formats = [ ".xlsx"];
const select_fields: FieldMap = {'part_number':["part number", "part no", "pn"], 'part_description':["part description", "part desc", "description", "desc"], 'part_quantity':["part quantity", "part qty", "quantity", "qty"], 
                'unit_prices':["part price", "unit price", "price"], 'unit_currency': ["currency", "curr", "cur"],
                'vendor_code':["vendor code", "vendor no", "vendor"], 'vendor_name':["vendor name", "vendor"], 
                'purchase_order':["purchase order", "po no", "related po"], 'order_type':["order type", "po type", "po"]};

const mandatory: string[] = ['part_number', 'part_description', 'part_quantity', 'vendor_code'];

const filter_fields: FieldMap = {'ot_filter': ['order_type'], 'curr_filter': ['unit_currency'], 'vc_filter': ['vendor_code'],
                                    'pn_filter': ['part_number']};

var acceptedExts = available_formats.join(',');
var dropzoneText = `${generateExtensionString(available_formats)} uzantılı bir dosya sürükle veya yüklemek için tıkla`;



///FLOW ////
const popupFilter = document.getElementById('popup-filter')!;
const closeFilter = document.getElementById('close-filter')!;

closeFilter.addEventListener('click', () => {
    popupFilter.style.display = 'none';
});

popupFilter.addEventListener("submit", async (e) => {
    e.preventDefault();
    
    try{

        let selectedSheet = document.getElementById('filter-sheet-table').getElementsByClassName('bg-success')[0].id;
        let selectedCol = (document.querySelector(`#filterpicker`) as HTMLInputElement).value;
        let selectedRadio = (document.querySelector('input[name="radio-filter"]:checked') as HTMLInputElement).id;
        
        const filterMap: FieldMap = { };
        filterMap['sheet_name'] = [selectedSheet];
        filterMap['column_name'] = [selectedCol];
        filterMap['is_include'] = [selectedRadio];

        let pn_col_main = ((document.querySelector(`#part_number`) as HTMLInputElement).value);

        let unfilteredVals = await window.versions.filter_from_table(filterMap, pn_col_main);
        unfilteredVals.forEach(function(val) {
            document.getElementById(`${val}`).setAttribute("selected", "selected");
        });

        $('.selectpicker').each(function() {
            (<any>$( this )).selectpicker('refresh');
        });

        setAnalysisMessage(AnalysisStage.SelectFields, {canAnalyze: true, message: "Total of " + unfilteredVals.length + " pn selected for analysis", infoClass:"text-info"});
    }
    catch{
        setAnalysisMessage(AnalysisStage.SelectFields, {canAnalyze: true, message: "Please fill all fields before submit", infoClass:"text-danger"});
    }
    
    popupFilter.style.display = 'none';

});

/*const chooseText = document.getElementById('choose-text');
let upload_name = "upload_link";
chooseText.innerHTML = `Press <span name="${upload_name}" class="link-primary" href="#" >boss</span> to start.${dropzoneText}`*/

setAnalysisMessage(AnalysisStage.Import);
// FLOW END //








//// FUNCTIONS /////

async function file_upload(filePath: string, xlsxType = 'main') {

    const sheets = await window.versions.xlsx_upload(xlsxType === 'main' ? 0 : 1, filePath);
    if (sheets == null)
    {
        return;
    }

    if(xlsxType == 'main'){
        draw_sheet_cols(sheets, 'sheet-table', mainSheetSelected);

    }else if(xlsxType == 'pn_filter'){
        draw_sheet_cols(sheets, 'filter-sheet-table', pnFilterSheetSelected);
    }
    else if(xlsxType == 'vc_filter'){
        draw_sheet_cols(sheets, 'filter-sheet-table', vendorFilterSheetSelected);
    }
}

function draw_sheet_cols(sheet_names: string[], table_name: string, onselect: (selected: string) => void){
    const sheetTable = document.getElementById(table_name);
    
    let sheetDivs = sheetTable.children;
    while (sheetDivs.length > 0) {
        sheetDivs[0].remove();
    }

    sheet_names.forEach(name => {
        const sheetDiv = document.createElement("div");
        sheetDiv.classList.add("col", "col-lg-4", "text-center", "vega-sheet-select");
        sheetDiv.textContent = `${name}`;
        sheetDiv.id = `${name}`;
        // add an onclick handler to the sheet div
        sheetDiv.onclick = async () => {
            let selected: string = null;

            if(sheetDiv.classList.contains("bg-success"))
            {
                sheetDiv.classList.remove("bg-success", "text-white");
                selected = null;
            }
            else
            {
                // set the selectedSheet variable to the current sheet number
                selected = name;
                // update the class of all sheet divs
                const sheetDivs = document.querySelectorAll("#" + table_name + " > .col");
                sheetDivs.forEach(div => {
                    div.classList.remove("bg-success", "text-white");
                });
                sheetDiv.classList.add("bg-success", "text-white");
                await onselect(selected);
            }
        };

        // append the sheet div to the sheet table
        sheetTable.appendChild(sheetDiv);
    });

    setAnalysisMessage(AnalysisStage.SelectSheet);
}

async function mainSheetSelected(selected: string){
    await draw_analysis_options(selected);
    analysisFieldChanged(selected);
}

async function pnFilterSheetSelected(selected: string){
    await draw_filter_by_table(selected, 'part_number');
    //analysisFieldChanged(selected);
}

async function vendorFilterSheetSelected(selected: string){
    await draw_filter_by_table(selected, 'vendor_code');
    //analysisFieldChanged(selected);
}

const draw_analysis_options = async (selectedSheet: string) => {
    const columns = await window.versions.get_columns_on_sheet(0, selectedSheet); // 0 for main excel table
    
    const fieldSelect = document.getElementById('analysis-field-selection');
    fieldSelect.innerHTML = "";

    const mandatoryDiv = document.createElement("div");
    mandatoryDiv.classList.add("parse-options");
    
    let mandatoryHeader = document.createElement("h4");
    mandatoryHeader.textContent = "Zorunlu Kolonlar";
    mandatoryDiv.appendChild(mandatoryHeader);

    let divider = document.createElement("hr");
    divider.classList.add("hr-vega");
    mandatoryDiv.appendChild(divider);

    const optionalDiv = document.createElement("div");
    optionalDiv.classList.add("parse-options");

    let optionalHeader = document.createElement("h4");
    optionalHeader.textContent = "Opsiyonel Kolonlar";
    optionalDiv.appendChild(optionalHeader);

    let divider2 = document.createElement("hr");
    divider2.classList.add("hr-vega");
    optionalDiv.appendChild(divider2);

    for (let key in select_fields) {

        let selectFieldHTML = '<div class="vega-select">';
        
        const text = key.replace('_', ' ').toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ') + ':';
        selectFieldHTML += '<label for="' + key + '">' + text + '</label>';
        
        selectFieldHTML += `<select class="selectpicker" onchange="analysisFieldChanged('${selectedSheet}')" style="float:right;" id="` + key + '" name="' + key + '">';

        let litteralDiff = 20;
        columns.forEach(c => {
            //const hardCodedName = field.replace('_', ' ');
            selectFieldHTML += '<option value="' + c + '"';
            const index = select_fields[key].findIndex((searchTerm: string) => c.toLocaleLowerCase().includes(searchTerm.toLocaleLowerCase()));
            if (index !== -1) {
                let currLitteralDiff = c.length - select_fields[key][index].length;
                if(currLitteralDiff < litteralDiff){
                    litteralDiff = currLitteralDiff;
                    selectFieldHTML += ' selected="selected"';
                }
            }
            selectFieldHTML += '>' + c + '</option>';
        });

        if(mandatory.includes(key) && 4 < litteralDiff){ //if mandatory field and not looks like any column
            selectFieldHTML += '<option value="" selected hidden="hidden">Please select</option>';
        }
        else if(!mandatory.includes(key)){
            let didSelected = 4 < litteralDiff ? "selected" : "";
            selectFieldHTML += '<option value="exclude" ' + didSelected + '>Exclude option</option>';
        }
        
        selectFieldHTML += '</select>';
        selectFieldHTML += '</div>'

        if(mandatory.includes(key)){
            mandatoryDiv.innerHTML += selectFieldHTML;
        }
        else{
            optionalDiv.innerHTML += selectFieldHTML;
        }

    }
    
    fieldSelect.appendChild(mandatoryDiv);
    //fieldSelect.appendChild(document.createElement("hr"))
    fieldSelect.appendChild(optionalDiv);
};

const draw_analysis_filters = async (selectedSheet: string) => {
    const analysisFilters = document.getElementById('analysis-field-filters');
    analysisFilters.innerHTML = "";

    const filterDiv = document.createElement("div");
    filterDiv.classList.add("filter-options");

    let filterHeader = document.createElement("h4");
    filterHeader.textContent = "Filtreler";
    filterDiv.appendChild(filterHeader);
    
    let divider = document.createElement("hr");
    divider.classList.add("hr-vega");
    filterDiv.appendChild(divider);

    let allSet = true;
    //document.createElement('select');

    for (let fieldKey in filter_fields) {
        
        let htmlFilter = '';
        let filterSelect = filter_fields[fieldKey][0];

        htmlFilter += '<div class="vega-select">';

        let optionSelected = (document.querySelector('#' + filterSelect + ' option:checked') as HTMLSelectElement).value;
        const text = fieldKey.replace('_', ' ').toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ') + ':';
        htmlFilter += '<label for="' + fieldKey + '" class="form-label select-label">' + text + '</label>';

        let disabled = (!optionSelected || optionSelected === 'exclude') ? 'disabled' : '';
        htmlFilter += '<select class="selectpicker" multiple data-live-search="true" style="float:right;" ' + disabled + ' id="' + fieldKey + '" name="' + fieldKey + '">';

        if (!disabled) {
            let columnValues = await window.versions.get_values_on_column(0, selectedSheet, optionSelected);//0 for main excel table
            let distinctValues = Array.from(new Set(columnValues));
            distinctValues.forEach(c => {
                //const hardCodedName = field.replace('_', ' ');
                htmlFilter += '<option value="' + c + '" id="' + c + '"';
                htmlFilter += '>' + c + '</option>';
            });
        }
        else if(mandatory.includes(filterSelect)){
            allSet = false;
        }
        
        htmlFilter += '</select>';
        htmlFilter += '</div>'

        filterDiv.innerHTML += htmlFilter;
    }
    
    if(allSet){
        let tableFilterButton = document.createElement("button");
        tableFilterButton.classList.add("btn", "btn-info");
        tableFilterButton.setAttribute("type", "button");
        tableFilterButton.setAttribute("onclick", "filter_from_sheet('pn')");
        tableFilterButton.setAttribute("id", "filter_pns");
        tableFilterButton.textContent = "Filter from PNs";

        let filterText = document.createElement("p");
        filterText.textContent = "Filter multiple part numbers from a loaded table. Exclude or include multiple Part Numbers";
        
        tableFilterButton.appendChild(filterText);
        filterDiv.appendChild(tableFilterButton);
    }

    analysisFilters.appendChild(filterDiv);

    /*let buttonHtml = `<div> \
        &emsp;<button class="btn btn-info" type="button" onclick="filter_from_sheet('pn')" id="filter_pns">Filter from PNs</button> \
        &emsp;<button class="btn btn-info" type="button" onclick="filter_from_sheet('vendor')" id="filter_vendors">Filter from Vendors</button> \
            </div>`;*/
    return allSet;
}

async function analysisFieldChanged(selectedSheet: string){
    var allSet = await draw_analysis_filters(selectedSheet);
    $('.selectpicker').each(function() {
        (<any>$( this )).selectpicker('refresh');
    });
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

    parts_map["sheet_name"] = [getSelectedMainTable()];

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
        setAnalysisMessage(AnalysisStage.SelectFields, {canAnalyze: true, message: "There is an error somewhere.. May be change select fields ?", infoClass:"text-danger"});
    }
}

const analysis_save = async () => {

    let selectedName = getSelectedMainTable();
    
    const saveResult = await window.versions.save_results(selectedName);
    console.log(saveResult);
    if(saveResult.length > 1){
        let analysis_text = "Saved to " + saveResult + " successfully";
        setAnalysisMessage(AnalysisStage.Exported, {message: analysis_text, infoClass:"text-success"});

    }
    else{
        let analysis_text = "Save unsuccessfull :(";
        setAnalysisMessage(AnalysisStage.Exported, {message: analysis_text, infoClass:"text-danger"});
    }
}


const draw_filter_by_table = async (selectedFilterSheet: string, filterType: string) => {
    const filterMessages = document.getElementById('filter_col_messsage');
    let filterColHtml = '';
    const filterCols = await window.versions.get_columns_on_sheet(1, selectedFilterSheet); // 1 for filter table

    filterColHtml += '<div>';
    const text = filterType.replace('_', ' ').toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
    filterColHtml += '<label for="' + filterType + '">' + text + ':</label>';
        
    filterColHtml += '<select class="filterpicker" id="filterpicker" name="filterpicker">';

    let litteralDiff = 20;
    filterCols.forEach(c => {
        //const hardCodedName = field.replace('_', ' ');
        filterColHtml += '<option value="' + c + '"';
        const index = select_fields[filterType].findIndex((searchTerm: string) => c.toLocaleLowerCase().includes(searchTerm.toLocaleLowerCase()));
        if (index !== -1) {
            let currLitteralDiff = c.length - select_fields[filterType][index].length;
            if(currLitteralDiff < litteralDiff){
                litteralDiff = currLitteralDiff;
                filterColHtml += ' selected="selected"';
            }
        }
        filterColHtml += '>' + c + '</option>';
    });

    if(4 < litteralDiff){
        filterColHtml += '<option value="" selected="selected" hidden="hidden">Please select</option>';
    }

    filterColHtml += '</select>';
    filterColHtml += '</div>';
    filterColHtml += '<div>';
    filterColHtml += '<div class="form-check"> \
                    <input class="form-check-input" type="radio" name="radio-filter" id="include" checked> \
                    <label class="form-check-label" for="include">Include all same ' + text + '</label>' + '\
                    </div>';

    filterColHtml += '<div class="form-check"> \
                    <input class="form-check-input" type="radio" name="radio-filter" id="exclude"> \
                    <label class="form-check-label" for="exclude">Exclude all same ' + text + '</label>' + '\
                    </div>';
    filterColHtml += '</div>';

    filterMessages.innerHTML = filterColHtml;

    (<any>$('.filterpicker')).selectpicker('refresh');
}

const filter_from_sheet = async (filterType: string) => {

    popupFilter.style.display = 'block';
    
    if (filterType === 'pn'){

    }
    else if(filterType === 'vendor'){
        
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

function getSelectedMainTable(){
    let sheetTable = document.getElementById('sheet-table');
    let selected = sheetTable.querySelector('.bg-success');
    return selected.textContent ?? "";
}

function setAnalysisMessage(currentStage: string, {message = '', infoClass = '', canAnalyze = false} = {}){

    const analysisMessages = document.getElementById('analysis-messages');

    let firstLine = ':`(';
    let secondClass = '';
    let secondLine = ':`((((';
    let firstButton = '';
    let secondButton = '';
    let thirdLine = '';

    switch (currentStage) {
        case AnalysisStage.Import:
            firstLine = 'Please select or drop a excel file to dropdown';
            secondClass = 'text-success';
            secondLine = '';
            break;
        case AnalysisStage.SelectSheet:
            firstLine = 'Select a table (from loaded excel) for analysis';
            secondClass = 'text-success';
            secondLine = '';
            break;
        case AnalysisStage.SelectFields:
            firstLine = 'Select the fields for analysis. Fields are set from the first row of loaded Excel table.';
            
            if (message != ''){
                secondLine = message;
                secondClass = infoClass;
            }
            else{
                secondLine = '!!! Check if the auto placed fields are correct !!!';
                secondClass = 'text-danger';
            }
            
            if (canAnalyze){
                thirdLine = '<span style="float: right;">All fields are selected. Select filters if you like. <br><span class="text-warning">!!Filters and optional colons are not tested in detail!!</span></span>';
                firstButton = `<button class="btn btn-success" type="button" onclick="parts_submit()" id="analyse_button">Analyse Parts</button>`;
            }
            break;
        case AnalysisStage.Analyzed:
            firstLine = 'You can now press save analysis to export Table';
            secondClass = 'text-success';
            secondLine = message;
            firstButton = `<button class="btn btn-success" type="button" onclick="parts_submit()" id="analyse_button">Analyse Parts</button>`;
            secondButton = '<button class="btn btn-secondary" onclick="analysis_save()" type="button">Save Analysis</button>';
            break;
        case AnalysisStage.Exported:
            if(infoClass.includes('success'))
            {
                firstLine = '';
            }
            //firstLine = 'Save is success!!';
            secondClass = infoClass;
            secondLine = message;
            firstButton = `<button class="btn btn-success" type="button" onclick="parts_submit()" id="analyse_button">Analyse Parts</button>`;
            secondButton = '<button class="btn btn-secondary" onclick="analysis_save()" type="button">Save Analysis</button>';
            break;
        default:
            firstLine = 'ERROR!!';
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