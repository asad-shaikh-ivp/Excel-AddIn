import * as React from 'react';
import { Spinner, SpinnerType } from 'office-ui-fabric-react';
import Header from './Header';
import { HeroListItem } from './HeroList';
import Progress from './Progress';
import StartPageBody from './StartPageBody';
import GetDataPageBody from './GetDataPageBody';
import SuccessPageBody from './SuccessPageBody';
import MnemonicListBody from './MnemonicListBody';
import OfficeAddinMessageBar from './OfficeAddinMessageBar';
import { logoutFromO365, signInO365 } from '../../utilities/office-apis-helpers';
import { loadTheme } from 'office-ui-fabric-react/lib/Styling';
loadTheme({
    palette: {
        themePrimary: '#10893e',
        themeLighterAlt: '#effdf4',
        themeLighter: '#dffbea',
        themeLight: '#bff7d5',
        themeTertiary: '#7aefa7',
        themeSecondary: '#14a94e',
        themeDarkAlt: '#0f7c39',
        themeDark: '#0c602c',
        themeDarker: '#094c23',
        neutralLighterAlt: '#f8f8f8',
        neutralLighter: '#f4f4f4',
        neutralLight: '#eaeaea',
        neutralQuaternaryAlt: '#dadada',
        neutralQuaternary: '#d0d0d0',
        neutralTertiaryAlt: '#c8c8c8',
        neutralTertiary: '#a6a6a6',
        neutralSecondaryAlt: '#767676',
        neutralSecondary: '#666666',
        neutralPrimary: '#333',
        neutralPrimaryAlt: '#3c3c3c',
        neutralDark: '#212121',
        black: '#000000',
        white: '#fff',
        red : '#f44336'
        //primaryBackground: '#fff',
        //primaryText: '#333'
    }
});
export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
    authStatus?: string;
    fileFetch?: string;
    headerMessage?: string;
    errorMessage?: string;
    userName?: string;

}

export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            authStatus: 'notLoggedIn',
            fileFetch: 'notFetched',
            headerMessage: '',
            errorMessage: '',
            userName:''
        };

        // Bind the methods that we want to pass to, and call in, a separate
        // module to this component. And rename setState to boundSetState
        // so code that passes boundSetState is more self-documenting.
        this.boundSetState = this.setState.bind(this);
        this.setToken = this.setToken.bind(this);
        this.displayError = this.displayError.bind(this);
        this.login = this.login.bind(this);
    }

    /*
        Properties
    */

    // The access token is not part of state because React is all about the
    // UI and the token is not used to affect the UI in any way.
    accessToken: string;
    userName: string;

    listItems: HeroListItem[] = [
    ];

    /*
        Methods
    */

    boundSetState: () => {};

    setToken = (accesstoken: string, userName?: string) => {
        this.accessToken = accesstoken;

        // // token_handling
        //console.log('access token set from react:', accesstoken);
        // console.log('at this time idtoken is :', localStorage.getItem('msal.idtoken') );

        OfficeRuntime.storage.removeItem('token').then(() => {
            OfficeRuntime.storage.setItem('token', accesstoken);
        });

        this.userName = userName !== null ? userName: '';
    }

    displayError = (error: string) => {
        this.setState({ errorMessage: error });
    }

    // Runs when the user clicks the X to close the message bar where
    // the error appears.
    errorDismissed = () => {
        this.setState({ errorMessage: '' });

        // If the error occured during a "in process" phase (logging in or getting files),
        // the action didn't complete, so return the UI to the preceding state/view.
        this.setState((prevState) => {
            if (prevState.authStatus === 'loginInProcess') {
                return { authStatus: 'notLoggedIn' };
            }
            else if (prevState.fileFetch === 'fetchInProcess') {
                return { fileFetch: 'notFetched' };
            }
            return null;
        });
    }

    login = async () => {
        await signInO365(this.boundSetState, this.setToken, this.displayError);
    }

    logout = async () => {
        await logoutFromO365(this.boundSetState, this.displayError);
    }

    setParentState = async() => {
        this.setState({  fileFetch:'notFetched', headerMessage: '' });

    }

    mnemonicList =async () => {


        this.setState({ fileFetch: 'displayMnemonic' });

        //@ts-ignore

        // OfficeRuntime.storage.getItem('token').then((authentication_token) => {


        //     //@ts-ignore
        //     let url = `${API_URL}/api/GetMnemonicList`;
        //     const requestHeaders: HeadersInit = new Headers();

        //     requestHeaders.append('Authorization', 'Bearer ' + authentication_token);
        //     requestHeaders.append('Content-Type', 'application/json');

        //     fetch(url, { method: 'GET'})
        //         .then((response) => {
        //             return response.json();
        //         })
        //         .then((mnemonics) => {
        //             //@ts-ignore
        //             let dialogUrl = `${ADDIN_URL}/Dialog/mnemonicList.html`
        //             Office.context.ui.displayDialogAsync(dialogUrl, { height: 80, width: 20 })

        //             window.localStorage.setItem('mnemonics', JSON.stringify(mnemonics))
        //             console.log(mnemonics)
        //         }).catch((error) => {
        //             console.log("error:" + error);
        //         });
        // });
    }




    uploadModel = async () => {
        this.setState({ fileFetch: 'fetchInProcess' });
        //@ts-ignore
        OfficeRuntime.storage.getItem('token').then((authentication_token) => {
            const requestHeaders: HeadersInit = new Headers();

            requestHeaders.append('Authorization', 'Bearer ' + authentication_token);
            requestHeaders.append('Content-Type', 'application/json');

            //@ts-ignore
            let url = `${API_URL}/api/GetMnemonicList`;


            fetch(url, { method: 'GET', 'headers': requestHeaders })
                .then((response) => {
                    if (!response.ok) {
                        throw Error(response.statusText);
                    }
                    return response.json();
                })
                .then((mnemonicsFromService) => {

            //read the active worksheet and get coy name
            Excel.run((context) => {

                const mnemonicsColumn = `A`;
                const excelMaxRowValue = `99999`; // max excel rows supported in online version
                let referenceDataValuesColumnIndex = 3; // adding new columns beside mnemonic will impact this index
                let numericValuesStartColumnIndex = 4;

                const yearMnemonic = `UD_C4_COMPANY_NAME`;
                const calendarDateMnemonic = `UD_C4_CALENDAR_DATE`;
                const timeDimensionMnemonic = `UD_C4_CURRENCY`;

                // logic specific to c4 company model
                // code will run on the active sheet on which user is present
                let currentWorkbook = context.workbook;
                //@ts-ignore
                context.workbook.save(Excel.SaveBehavior.save);
                let companyModelWorksheet = context.workbook.worksheets.getActiveWorksheet();
                let mnemonics = companyModelWorksheet.getRange(`${mnemonicsColumn}1:${mnemonicsColumn}${excelMaxRowValue}`);
                mnemonics.load('cellCount, values, columnIndex, rowIndex, text');
                currentWorkbook.load('isDirty');
                // placeholder for the sheet values
                // this is later converted to the request object
                //@ts-ignore
                let modelData = [];

                let usedRange = companyModelWorksheet.getUsedRange(true);
                usedRange.load('address, rowCount, columnCount');

                return context.sync()
                    .then(() => {
                        // here now since we get the used range value
                        // we are able to determine maximum rows available in the sheet
                        //@ts-ignore
                        excelMaxRowValue = usedRange.rowCount;
                        //@ts-ignore
                        let excelLastCol = usedRange.columnCount;

                        // since the object is returning values as array,
                        // we need to get the relevant address by index only
                        for (let ii = 0; ii < mnemonics.values.length; ii++) {
                            if (mnemonics.values[ii].toString() !== '') {
                                modelData.push({
                                    value: mnemonics.values[ii][0],
                                    cellIndex: ii,
                                    rowIndex: mnemonics.rowIndex,
                                    index: ii
                                });
                            }
                        }

                        // validate if the required basic mnemonics are present in the open excel sheet
                        // to be refactored into isSystemMnemonics fetched from server
                        // the following code only checks for existence of mnemonics in first column, not values
                        const mandatoryRequiredTags =
                            [`UD_C4_FY_START_MONTH`,
                                `UD_C4_COMPANY_ID`,
                                `UD_C4_COMPANY_NAME`];
                        //@ts-ignore
                        let errorCells = [];
                        //@ts-ignore
                        let invalidMnemonics = [];

                        let isRequiredDataPresent = true;
                        mandatoryRequiredTags.forEach(requiredTag => {
                            isRequiredDataPresent = isRequiredDataPresent && (modelData.find(m => m.value === requiredTag) !== undefined);
                        });

                        if (!isRequiredDataPresent) {
                            //@ts-ignore
                            Office.context.ui.displayDialogAsync(`${ADDIN_URL}/invalidModel.html`, { height: 30, width:40 , displayInIframe: true });
                            this.setState({ fileFetch: 'notFetched', headerMessage: '' });
                            return null;
                        }

                        // loading row-wise data against each mnemonic
                        // till the end of the last column
                        for (let jj = 0; jj < modelData.length; jj++) {
                            modelData[jj].rowReference = companyModelWorksheet.getRangeByIndexes(modelData[jj].index, 0, 1, excelLastCol);
                            context.load(modelData[jj].rowReference, 'cellCount, values, columnIndex, rowIndex, text');
                        }

                        return context.sync()
                            .then(() => {

                            // now to check if col b  and c are present for
                            // to upsert the mnemonic entries
                            let isUpsertSheet = true;
                            for (let jj = 0; jj < modelData.length; jj++) {
                                for (let kk = 0; kk < modelData[jj].rowReference.values.length; kk++) {
                                    isUpsertSheet = isUpsertSheet && ['Text', 'Number','Ratio'].includes (modelData[jj].rowReference.values[kk][1]);
                                }
                            }

                            console.log("isUpsertSheet :" + isUpsertSheet);
                            if(isUpsertSheet){
                                // upsert process called
                                console.log("Mnemonic upsert process");
                                //modelData[jj].isReferenceDataType  = mnemonic.mnemonic_type_name == "Text";
                                for (let jj = 0; jj < modelData.length; jj++) {
                                    for (let kk = 0; kk < modelData[jj].rowReference.values.length; kk++) {
                                    modelData[jj].isReferenceDataType = (modelData[jj].rowReference.values[kk][1] === 'Text');
                                    }
                                }
                            }else{

                                referenceDataValuesColumnIndex = 1;
                                numericValuesStartColumnIndex = 2;

                                // check here if any of the mnemonics
                                //which are not present in the db
                                // if unknown mnemonic is encountered add it
                                // to list for highlighting
                                for (let jj = 0; jj < modelData.length; jj++) {
                                    for (let kk = 0; kk < modelData[jj].rowReference.values.length; kk++) {

                                        let mnemonic = mnemonicsFromService.mnemonicList.find(m=>m.mnemonicName.trim().toUpperCase() === modelData[jj].value.trim().toUpperCase());
                                        if(mnemonic == null){
                                            // push to invalid_mnemonics list for highlighting
                                            invalidMnemonics.push({ rowIndex: modelData[jj].index, index: 0 });
                                        }else{
                                            modelData[jj].isReferenceDataType  = mnemonic.mnemonicTypeName == "Text";
                                        }
                                    }
                                }
                            }

                                for (let jj = 0; jj < modelData.length; jj++) {

                                    // mapping all rows to relevant data
                                    modelData[jj].filteredData = [];

                                    // Verify if only one dimension cell structure will suffice here
                                    for (let kk = 0; kk < modelData[jj].rowReference.values.length; kk++) {

                                        for (let ll = 0; ll < modelData[jj].rowReference.values[kk].length; ll++) {

                                            //modelData[jj].isReferenceDataType = (modelData[jj].rowReference.values[kk][1] === 'Text');

                                            if (modelData[jj].rowReference.values[kk][ll] !== '') {
                                                modelData[jj].filteredData.push({
                                                    value: modelData[jj].rowReference.values[kk][ll].toString(),
                                                    text: modelData[jj].rowReference.text[kk][ll],
                                                    index: ll,
                                                    rowIndex: modelData[jj].rowReference.rowIndex
                                                });
                                            }
                                        }
                                    }
                                }

                                return context.sync()
                                    .then(() => {
                                       
                                        isRequiredDataPresent = true;
                                        mandatoryRequiredTags.forEach(requiredTag => {
                                            //isRequiredDataPresent = isRequiredDataPresent && (modelData.find(m => m.value === requiredTag) !== undefined);
                                            let reqdTag = modelData.find(m => m.value === requiredTag);
                                            let refValue =reqdTag.filteredData.find(r=>r.index == referenceDataValuesColumnIndex );
                                            if(refValue === undefined || refValue=== null || refValue === ''){
                                                isRequiredDataPresent = false;
                                            }
                                        });

                                        if (!isRequiredDataPresent) {
                                            //@ts-ignore
                                            Office.context.ui.displayDialogAsync(`${ADDIN_URL}/invalidModel.html`, { height: 30, width: 40, displayInIframe: true });
                                            this.setState({ fileFetch: 'notFetched', headerMessage: '' });
                                            return null;
                                        }
                                        //preparing a list of time dimensions
                                        let yearRow = modelData.find(m => m.value === yearMnemonic);
                                        let dimensionRow = modelData.find(m => m.value === timeDimensionMnemonic);
                                        let asOfDateRow = modelData.find(m => m.value === calendarDateMnemonic);

                                        let asOfDateRowIndex = asOfDateRow.cellIndex;


                                        const yearFormat = RegExp(/^(\d{4})|(\d{4}$)/)
                                        const dimensionFormat = RegExp(/^LTM$|^YTD$|^Q\d$|^Y$|^Q$|^H\d$/i);
                                        let timedimensions = [];


                                        for(let ii = numericValuesStartColumnIndex; ii < yearRow.filteredData.length; ii++){
                                            let yearVal='', dimensionVal='', yearMatchesValues
                                            let dimensionY =''
                                            let dimensionD = ''
                                            let validDimensionYearRow = [], validDimensionDimRow = [],validYearYRow =[], validYearDRow = []

                                            //Checking if dimension is valid in either of the 2 rows
                                            if(yearRow.filteredData[ii] !== null && yearRow.filteredData[ii].value){
                                             validDimensionYearRow = (String(yearRow.filteredData[ii].value)).match(dimensionFormat)
                                            }

                                            if(dimensionRow.filteredData[ii] && dimensionRow.filteredData[ii].value){
                                                validDimensionDimRow = (String(dimensionRow.filteredData[ii].value)).match(dimensionFormat)
                                            }

                                            //checking if year is valid in either of the two 2rows
                                            if(validDimensionYearRow !== null && validDimensionYearRow.length > 0){
                                                dimensionY = validDimensionYearRow[0].toUpperCase()
                                            }
                                            if(validDimensionDimRow !== null && validDimensionDimRow.length > 0)
                                            {
                                                dimensionD = validDimensionDimRow[0].toUpperCase()
                                            }
                                            if(dimensionRow.filteredData[ii] && dimensionRow.filteredData[ii].value){
                                             validYearYRow = (String(yearRow.filteredData[ii].value)).match(yearFormat)
                                            }
                                            //typeof(yearRow.filteredData[ii]) !== "undefined"
                                            if(dimensionRow.filteredData[ii] && dimensionRow.filteredData[ii].value){
                                             validYearDRow = (String(dimensionRow.filteredData[ii].value)).match(yearFormat)
                                            }
                                            if (validYearYRow !== null && validYearYRow.length > 0) {
                                                yearVal = validYearYRow[0]
                                            } else if (validYearDRow !== null && validYearDRow.length > 0) {
                                                yearVal = validYearDRow[0]
                                            } else {
                                                continue;
                                            }

                                            if(dimensionY.length > 0 && dimensionD.length > 0){
                                                if(dimensionY === "LTM" || dimensionY === "YTD"){
                                                    dimensionVal = dimensionY
                                                } else{
                                                    dimensionVal = dimensionD
                                                }
                                            } else if (dimensionY.length > 0 && dimensionD.length === 0){
                                                dimensionVal = dimensionY
                                            } else if(dimensionY.length === 0 && dimensionD.length > 0){
                                                dimensionVal = dimensionD
                                            }

                                            let asOfDate = asOfDateRow.filteredData.find(d => d.index === yearRow.filteredData[ii].index);
                                            let asOfDateVal = '';
                                            if (asOfDate) {
                                                 asOfDateVal = asOfDate.text;
                                             }

                                             let asOfDateValParsed = Date.parse(asOfDateVal);
                                             if (isNaN(asOfDateValParsed) || asOfDateVal === '') {
                                                 errorCells.push({ rowIndex: asOfDateRowIndex, index: yearRow.filteredData[ii].index });
                                                 continue;
                                             }

                                             let date = new Date(asOfDateValParsed);
                                             if(yearVal.length > 0 && dimensionVal.length > 0){
                                                timedimensions.push({
                                                    index: yearRow.filteredData[ii].index,
                                                    year: yearVal,
                                                    timeDimension: dimensionVal,
                                                    asOfDate: `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`
                                                });
                                            }

                                         }

                                         console.log(timedimensions)
                                         // hence preparing the request object as per time dimensions
                                        let requestData = {
                                            mnemonicList: [],
                                            isMnemonicUpsertModel : isUpsertSheet
                                        };

                                        //@ts-ignore
                                        const numericFormat = RegExp(/-?[\d.]+(?:e-?\d+)?/);


                                        for (let ii = 0; ii < modelData.length; ii++) {
                                            let MnemonicDataType = '';
                                            let MnemonicCalculationType = '';

                                            // if in case this is an upsert sheet,
                                            // we need to pass the calc_type and data_type
                                            // to the server for inserting/updating mnemonics
                                            if(modelData[ii].filteredData.length >= 2){
                                                if(modelData[ii].filteredData[1] != null && modelData[ii].filteredData[1].hasOwnProperty("value"))
                                                    MnemonicDataType = modelData[ii].filteredData[1].value;
                                                if(modelData[ii].filteredData[2] != null && modelData[ii].filteredData[2].hasOwnProperty("value"))
                                                    MnemonicCalculationType = modelData[ii].filteredData[2].value;
                                            }
                                            let mnemonicEntry = {
                                                index: ii,
                                                name: modelData[ii].value,
                                                value: '',
                                                dimensionValues: [],
                                                MnemonicDataType: MnemonicDataType,
                                                MnemonicCalculationType: MnemonicCalculationType
                                            };

                                            if (modelData[ii].isReferenceDataType && modelData[ii].value !== 'UD_C4_CALENDAR_DATE') {
                                                mnemonicEntry.value = modelData[ii].filteredData[referenceDataValuesColumnIndex].value;
                                            }
                                            else if (modelData[ii].value === 'UD_C4_CURRENCY') {
                                                for (let jj = numericValuesStartColumnIndex; jj < modelData[ii].filteredData.length; jj++) {
                                                    mnemonicEntry.value = modelData[ii].filteredData[jj].value;
                                                }
                                            } else {
                                                // starting from third column, since dimensions data starts at column 3
                                                for (let jj = numericValuesStartColumnIndex; jj < modelData[ii].filteredData.length; jj++) {
                                                    let dimention = timedimensions.find(t => t.index === modelData[ii].filteredData[jj].index);
                                                    if (dimention != null) {

                                                        if (numericFormat.test(modelData[ii].filteredData[jj].value)|| modelData[ii].value === 'UD_C4_CALENDAR_DATE') {
                                                            //     if(companyModelWorksheet.getCell(ii, jj).format.fill.color === 'orange'){
                                                            //         companyModelWorksheet.getCell(ii, jj).format.fill.color = 'white';
                                                            //     }
                                                            if (modelData[ii].value === 'UD_C4_CALENDAR_DATE') {
                                                                mnemonicEntry.dimensionValues.push({

                                                                    value: modelData[ii].filteredData[jj].text,
                                                                    asOfDate: dimention.asOfDate,
                                                                    //asOfDateText: dimention.asOfDateText,
                                                                    timeDimension: dimention.timeDimension,
                                                                    year: dimention.year
                                                                });

                                                            } else if ((modelData[ii].value !== 'UD_C4_CALENDAR_DATE' && !isNaN(modelData[ii].filteredData[jj].value))) {

                                                                mnemonicEntry.dimensionValues.push({

                                                                    value: modelData[ii].filteredData[jj].value,
                                                                    asOfDate: dimention.asOfDate,
                                                                    //asOfDateText: dimention.asOfDateText,
                                                                    timeDimension: dimention.timeDimension,
                                                                    year: dimention.year
                                                                });
                                                            }

                                                        }
                                                        else {

                                                            if (modelData[ii].filteredData[jj].value || errorCells.length > -1) {
                                                                console.dir(errorCells)
                                                                errorCells.push(modelData[ii].filteredData[jj]);
                                                                //window.localStorage.setItem('errorcells',JSON.stringify(errorCells))
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            requestData.mnemonicList.push(mnemonicEntry);
                                        }

                                        if (errorCells.length > 0) {
                                            console.log(errorCells)
                                            //todo: see if this mechanism can be replaced by excel's native error handling mechanism
                                            errorCells.forEach(cell => {
                                                companyModelWorksheet.getCell(cell.rowIndex, cell.index).format.fill.color = 'orange';
                                            });
                                            //@ts-ignore
                                            let dialogUrl = `${ADDIN_URL}/ErrorDialog/dialog.html`;
                                            Office.context.ui.displayDialogAsync(dialogUrl, { height: 25, width: 40 })
                                            this.setState({ fileFetch: 'notFetched', headerMessage: '' });
                                            return;
                                        }

                                        if (invalidMnemonics.length > 0) {

                                            //todo: see if this mechanism can be replaced by excel's native error handling mechanism
                                            invalidMnemonics.forEach(cell => {
                                                companyModelWorksheet.getCell(cell.rowIndex, cell.index).format.fill.color = 'orange';
                                            });
                                            //@ts-ignore
                                            let dialogUrl = `${ADDIN_URL}/ErrorDialog/invalidMnemonicsInModel.html`;
                                            Office.context.ui.displayDialogAsync(dialogUrl, { height: 25, width: 40 })
                                            this.setState({ fileFetch: 'fetched', headerMessage: '' });
                                            return;
                                        }
                                        //send the parsed data to the server
                                        //@ts-ignore
                                        let url = `${API_URL}/api/ParseExcelModelDataWithdimensions`;
                                        fetch(url, { method: 'POST', 'headers': requestHeaders, body: JSON.stringify(requestData) })
                                            .then((response) => {
                                                console.log(requestData)
                                                // server throws a non-200 message if there are any business exceptions
                                                if (!response.ok) {
                                                   
                                                    try {
                                                        return response.json();
                                                    } catch (e) {
                                                        return response.text();
                                                    }
                                                  }
                                                return response.json();
                                            })
                                            .then((mnemonics) => {

                                                //@ts-ignore
                                                let dialogUrl = `${ADDIN_URL}/Dialog/dialog.html`

                                                if(mnemonics === "#N/A Issuer/Asset Not Found"){
                                                    //@ts-ignore
                                                    dialogUrl = `${ADDIN_URL}/ErrorDialog/invalidIdentifier.html`
                                                    Office.context.ui.displayDialogAsync(dialogUrl, { height: 20, width: 30});
                                                    this.setState({ fileFetch: 'notFetched', headerMessage: '' });
                                                    return;
                                                }else if(mnemonics === "Invalid Mnemonic"){
                                                    //@ts-ignore
                                                    dialogUrl = `${ADDIN_URL}/ErrorDialog/excelRunError.html`;
                                                    Office.context.ui.displayDialogAsync(dialogUrl,{ height: 20, width: 30});
                                                    var errorMessage = JSON.stringify({
                                                        header: "Error uploading data",
                                                        text :"Invalid Mnemonic detected in the system. Please contact Administrator."
                                                    })
                                                    window.localStorage.setItem('error', errorMessage)
                                                      this.setState({ fileFetch: 'notFetched', headerMessage: '' });
                                                    return;
                                                }else{
                                                    Office.context.ui.displayDialogAsync(dialogUrl, { height: 600, width: 800 });
                                                }

                                                window.localStorage.setItem('mnemonics', JSON.stringify(mnemonics))
                                                window.localStorage.setItem('requestData', JSON.stringify(requestData))
                                                this.setState({ fileFetch: 'fetched', headerMessage: '' });

                                            }).catch((requestError) => {

                                                console.log("error from service:" + requestError);
                                                console.log(requestError)
                                                console.log(Error)
                                                
                                                //@ts-ignore
                                                let dialogUrl = `${ADDIN_URL}/ErrorDialog/fetchError.html`
                                                Office.context.ui.displayDialogAsync(dialogUrl, { height: 20, width: 30 });

                                                if(String(requestError) === "#N/A Issuer/Asset Not Found"){
                                                    //@ts-ignore
                                                    dialogUrl = `${ADDIN_URL}/ErrorDialog/invalidIdentifier.html`
                                                }

                                                Office.context.ui.displayDialogAsync(dialogUrl, { height: 20, width: 30 });
                                                this.setState({ fileFetch: 'notFetched', headerMessage: '' });
                                            });
                                    });
                            });

                    });
            }).catch((err) => {

                console.log(err)
                //todo : handle error in a new dialog
                //@ts-ignore
                let dialogUrl = `${ADDIN_URL}/ErrorDialog/excelRunError.html`;
                Office.context.ui.displayDialogAsync(dialogUrl,{ height: 20, width: 30});
                var errorMessage = JSON.stringify({
                    header: " Error in uploading the data",
                    text :"The cell with mnemonic value as text cannot be empty"
                })
                        window.localStorage.setItem('error', errorMessage)

                this.setState({ fileFetch: 'notFetched', headerMessage: '' });
            });
        }).catch((err) => {
            console.log(err)
            //todo : handle error in a new dialog
            //@ts-ignore
            let dialogUrl = `${ADDIN_URL}/ErrorDialog/excelRunError.html`;
            Office.context.ui.displayDialogAsync(dialogUrl,{ height: 20, width: 30});
            var errorMessage = JSON.stringify({
                header: "Error uploading data",
                text :"An error occurred while uploading model. Please contact Administrator."
            })
                    window.localStorage.setItem('error', errorMessage)

            this.setState({ authStatus: 'notLoggedIn',fileFetch: 'notFetched', headerMessage: '' });
        });
        });
    }

    deleteModel = async () => {

        //@ts-ignore
        OfficeRuntime.storage.getItem('token').then((authentication_token) => {
            const requestHeaders: HeadersInit = new Headers();

            requestHeaders.append('Authorization', 'Bearer ' + authentication_token);
            requestHeaders.append('Content-Type', 'application/json');


              //@ts-ignore
              let url = `${API_URL}/api/GetMnemonicList`;
              fetch(url, { method: 'GET', 'headers': requestHeaders })
                  .then((response) => {
                      if (!response.ok) {
                          throw Error(response.statusText);
                      }
                      return response.json();
                  })
                  .then((mnemonicsFromService) => {

                    Excel.run((context) => {

                        const mnemonicsColumn = `A`;
                        const excelMaxRowValue = `99999`; // max excel rows supported in online version
                        let referenceDataValuesColumnIndex = 3; // adding new columns beside mnemonic will impact this index
                        let numericValuesStartColumnIndex = 4;
        
                        const yearMnemonic = `UD_C4_COMPANY_NAME`;
                        const calendarDateMnemonic = `UD_C4_CALENDAR_DATE`;
                        const timeDimensionMnemonic = `UD_C4_CURRENCY`;
        
                        // logic specific to c4 company model
                        // code will run on the active sheet on which user is present
                        let currentWorkbook = context.workbook;
                        //@ts-ignore
                        context.workbook.save(Excel.SaveBehavior.save);
                        let companyModelWorksheet = context.workbook.worksheets.getActiveWorksheet();
                        let mnemonics = companyModelWorksheet.getRange(`${mnemonicsColumn}1:${mnemonicsColumn}${excelMaxRowValue}`);
                        mnemonics.load('cellCount, values, columnIndex, rowIndex, text');
                        currentWorkbook.load('isDirty');
                        // placeholder for the sheet values
                        // this is later converted to the request object
                        //@ts-ignore
                        let modelData = [];
        
                        let usedRange = companyModelWorksheet.getUsedRange(true);
                        usedRange.load('address, rowCount, columnCount');
        
                        return context.sync()
                            .then(() => {
                                // here now since we get the used range value
                                // we are able to determine maximum rows available in the sheet
                                //@ts-ignore
                                excelMaxRowValue = usedRange.rowCount;
                                //@ts-ignore
                                let excelLastCol = usedRange.columnCount;
        
                                // since the object is returning values as array,
                                // we need to get the relevant address by index only
                                for (let ii = 0; ii < mnemonics.values.length; ii++) {
                                    if (mnemonics.values[ii].toString() !== '') {
                                        modelData.push({
                                            value: mnemonics.values[ii][0],
                                            cellIndex: ii,
                                            rowIndex: mnemonics.rowIndex,
                                            index: ii
                                        });
                                    }
                                }
        
                                // validate if the required basic mnemonics are present in the open excel sheet
                                // to be refactored into isSystemMnemonics fetched from server
                                // the following code only checks for existence of mnemonics in first column, not values
                                const mandatoryRequiredTags =
                                    [`UD_C4_FY_START_MONTH`,
                                        `UD_C4_COMPANY_ID`,
                                        `UD_C4_COMPANY_NAME`];
                                //@ts-ignore
                                let errorCells = [];
                                //@ts-ignore
                                let invalidMnemonics = [];
        
                                let isRequiredDataPresent = true;
                                mandatoryRequiredTags.forEach(requiredTag => {
                                    isRequiredDataPresent = isRequiredDataPresent && (modelData.find(m => m.value === requiredTag) !== undefined);
                                });
        
                                if (!isRequiredDataPresent) {
                                    //@ts-ignore
                                    Office.context.ui.displayDialogAsync(`${ADDIN_URL}/invalidModel.html`, { height: 30, width:40 , displayInIframe: true });
                                    this.setState({ fileFetch: 'notFetched', headerMessage: '' });
                                    return null;
                                }
        
                                // loading row-wise data against each mnemonic
                                // till the end of the last column
                                for (let jj = 0; jj < modelData.length; jj++) {
                                    modelData[jj].rowReference = companyModelWorksheet.getRangeByIndexes(modelData[jj].index, 0, 1, excelLastCol);
                                    context.load(modelData[jj].rowReference, 'cellCount, values, columnIndex, rowIndex, text');
                                }
        
                                return context.sync()
                                .then(() => {
        
                                    // now to check if col b  and c are present for
                                    // to upsert the mnemonic entries
                                    let isUpsertSheet = true;
                                    for (let jj = 0; jj < modelData.length; jj++) {
                                        for (let kk = 0; kk < modelData[jj].rowReference.values.length; kk++) {
                                            isUpsertSheet = isUpsertSheet && ['Text', 'Number','Ratio'].includes (modelData[jj].rowReference.values[kk][1]);
                                        }
                                    }
        
                                    console.log("isUpsertSheet :" + isUpsertSheet);
                                    if(isUpsertSheet){
                                        // upsert process called
                                        console.log("Mnemonic upsert process");
                                        //modelData[jj].isReferenceDataType  = mnemonic.mnemonic_type_name == "Text";
                                        for (let jj = 0; jj < modelData.length; jj++) {
                                            for (let kk = 0; kk < modelData[jj].rowReference.values.length; kk++) {
                                            modelData[jj].isReferenceDataType = (modelData[jj].rowReference.values[kk][1] === 'Text');
                                            }
                                        }
                                    }else{
        
                                        referenceDataValuesColumnIndex = 1;
                                        numericValuesStartColumnIndex = 2;
        
                                        // check here if any of the mnemonics
                                        //which are not present in the db
                                        // if unknown mnemonic is encountered add it
                                        // to list for highlighting
                                        for (let jj = 0; jj < modelData.length; jj++) {
                                            for (let kk = 0; kk < modelData[jj].rowReference.values.length; kk++) {
        
                                                let mnemonic = mnemonicsFromService.mnemonicList.find(m=>m.mnemonicName.trim().toUpperCase() === modelData[jj].value.trim().toUpperCase());
                                                if(mnemonic == null){
                                                    // push to invalid_mnemonics list for highlighting
                                                    invalidMnemonics.push({ rowIndex: modelData[jj].index, index: 0 });
                                                }else{
                                                    modelData[jj].isReferenceDataType  = mnemonic.mnemonicTypeName == "Text";
                                                }
                                            }
                                        }
                                    }
        
                                        for (let jj = 0; jj < modelData.length; jj++) {
        
                                            // mapping all rows to relevant data
                                            modelData[jj].filteredData = [];
        
                                            // Verify if only one dimension cell structure will suffice here
                                            for (let kk = 0; kk < modelData[jj].rowReference.values.length; kk++) {
        
                                                for (let ll = 0; ll < modelData[jj].rowReference.values[kk].length; ll++) {
        
                                                    //modelData[jj].isReferenceDataType = (modelData[jj].rowReference.values[kk][1] === 'Text');
        
                                                    if (modelData[jj].rowReference.values[kk][ll] !== '') {
                                                        modelData[jj].filteredData.push({
                                                            value: modelData[jj].rowReference.values[kk][ll].toString(),
                                                            text: modelData[jj].rowReference.text[kk][ll],
                                                            index: ll,
                                                            rowIndex: modelData[jj].rowReference.rowIndex
                                                        });
                                                    }
                                                }
                                            }
                                        }

                                        let requestData =[]
                                        
                                        for (let ii = 0; ii < modelData.length; ii++) {
                                            if (modelData[ii].isReferenceDataType && modelData[ii].value === 'UD_C4_COMPANY_ID') {
                                             let mnemonicEntry = {
                                                index: ii,
                                                name: modelData[ii].value,
                                                value: modelData[ii].filteredData[referenceDataValuesColumnIndex].value,
                                                dimensionValues: []
                                            };
                                            requestData.push(mnemonicEntry)
                                            var issuer = requestData[0].value;
                                            break;
                                            }
                                        }
                                    
                                        //@ts-ignore
                                        let url = `${API_URL}/api/GetIssuerDetails?modelPrimaryIdentifier=${issuer}`;
                                        fetch(url, { method: 'GET', 'headers': requestHeaders })
                                            .then((response) => {
                                                console.log(requestData)
                                                // server throws a non-200 message if there are any business exceptions
                                                if (!response.ok) {
                                                   
                                                    try {
                                                        return response.json();
                                                    } catch (e) {
                                                        return response.text();
                                                    }
                                                  }
                                                return response.json();
                                            })
                                            .then((mnemonics) => {
                                                
                                                window.localStorage.setItem('companyID', JSON.stringify(requestData))
                                                window.localStorage.setItem('issuerDetails', JSON.stringify(mnemonics))
                                                var dialog;
                                                var deleted = false
                                                //@ts-ignore
                                                let dialogUrl = `${ADDIN_URL}/Dialog/deleteAI.html`;
                                                Office.context.ui.displayDialogAsync(dialogUrl,{height:45, width:40},function (asyncResult){
                                                    dialog = asyncResult.value;
                                                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                                                    });
                            
                                                function processMessage(arg) {
                                                  
                                                    var messageFromDialog = JSON.parse(arg.message);
                                                    if(messageFromDialog.messageType === "Yes"){
                                                    deleted = true
                                                        //@ts-ignore
                                                    let url = `${API_URL}/api/RemoveIssuer?modelPrimaryIdentifier=${issuer}`;
                                                    fetch(url, { method: 'GET', 'headers': requestHeaders })
                                                        .then((response) => {
                                                            console.log(requestData)
                                                            // server throws a non-200 message if there are any business exceptions
                                                            if (!response.ok) {
                                                            
                                                                try {
                                                                    return response.json();
                                                                } catch (e) {
                                                                    return response.text();
                                                                }
                                                            }

                                                            return response.json

                                                        }).then((response)=> {
                                                                //@ts-ignore
                                                                let dialogUrl = `${ADDIN_URL}/Dialog/success.html`;
                                                                Office.context.ui.displayDialogAsync(dialogUrl,{ height: 20, width: 30});
                                                        }).catch((err) => {
                                                            console.log(err)
                                                            //todo : handle error in a new dialog
                                                            //@ts-ignore
                                                            let dialogUrl = `${ADDIN_URL}/ErrorDialog/excelRunError.html`;
                                                            Office.context.ui.displayDialogAsync(dialogUrl,{ height: 20, width: 30});
                                                            var errorMessage = JSON.stringify({
                                                                header: "Error deleting issuer",
                                                                text :"An error occurred while model model. Please contact Administrator."
                                                            })
                                                                    window.localStorage.setItem('error', errorMessage)
                                                
                                                            this.setState({ authStatus: 'notLoggedIn',fileFetch: 'notFetched', headerMessage: '' });
                                                        });
                                                        dialog.close();
                                                    }
                                                    else {
                                                        dialog.close();
                                                    }
                                                   
                                                }


                                             

                                                if(mnemonics === "#N/A Issuer/Asset Not Found"){
                                                    //@ts-ignore
                                                    dialogUrl = `${ADDIN_URL}/ErrorDialog/invalidIdentifier.html`
                                                    Office.context.ui.displayDialogAsync(dialogUrl, { height: 20, width: 30});
                                                    this.setState({ fileFetch: 'notFetched', headerMessage: '' });
                                                    return;
                                                }

                            
                                                var errorMessage = JSON.stringify({
                                                    header: "Are you sure you want to delete the data for the following Issuer?",
                                                    text :"The data for the following Issuer will be deleted from the RMS"
                                                })
                                                        window.localStorage.setItem('error', errorMessage)
                            
                            
                                                this.setState({ fileFetch: 'notFetched', headerMessage: '' });
                                            })
                                 })
                         })

                    })

                   

            });

        })
    }


    render() {
        const { title, isOfficeInitialized } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo='assets/cap4_logo_white-80.png'
                    message='Please load your add-in in excel to use the app.'
                />
            );
        }

        // Set the body of the page based on where the user is in the workflow.
        let body;

        if (this.state.authStatus === 'notLoggedIn') {
            body = (<StartPageBody login={this.login} listItems={this.listItems} />);
        }
        else if (this.state.authStatus === 'loginInProcess') {
            body = (<Spinner className='spinner' type={SpinnerType.large} label='Please sign-in on the pop-up window.' />);
        }
        else if(this.state.fileFetch === "displayMnemonic"){

            body = (<MnemonicListBody  setParentState = {this.setParentState} />);
        }
        else {
            if (this.state.fileFetch === 'notFetched') {
                body = (<GetDataPageBody logout={this.logout} uploadModel={this.uploadModel} deleteModel = {this.deleteModel} mnemonicList = {this.mnemonicList}/>);
            }
            else if (this.state.fileFetch === 'fetchInProcess') {
                body = (<Spinner className='spinner' type={SpinnerType.large} label='Upload in progress' />);
            }
            else {
                body = (<SuccessPageBody logout={this.logout} uploadModel={this.uploadModel} deleteModel = {this.deleteModel} mnemonicList = {this.mnemonicList}/>);
            }


            // //// token_handling
            // //console.log('access token set :', accesstoken);
            // console.log('at this time idtoken is :', localStorage.getItem("msal.idtoken") );

            // OfficeRuntime.storage.removeItem('token').then(() => {
            //     OfficeRuntime.storage.setItem('token', localStorage.getItem("msal.idtoken"));
            // });
        }

        return (
            <div>
                {this.state.errorMessage ?
                    (<OfficeAddinMessageBar onDismiss={this.errorDismissed} message={this.state.errorMessage + ' '} />)
                    : null}

                <div className='ms-welcome'>
                    <Header logo='assets/cap4-title-large.png' title={this.props.title} message={this.state.headerMessage} userName={this.state.userName}/>
                    {body}
                </div>
            </div>
        );
    }
}
