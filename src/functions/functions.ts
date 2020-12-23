let isValidPeriod = (period) => {
  //console.log("period :"+period);
  let absoluteMnemonics = ["Y", "Q1", "Q2", "Q3", "Q4", "H1", "H2"];
  let relativeMnemonics = ["FY", "LTM", "YTD", "Q", "LASTAVAILABLEQ"];

  if (absoluteMnemonics.some(m => period.toUpperCase() == m)
||relativeMnemonics.some(m => period.toUpperCase() == m)) {
    return true
  } else if (relativeMnemonics.some(m => period.toUpperCase().startsWith(m))) {
    let relMnemonic = relativeMnemonics.find(m => period.toUpperCase().startsWith(m));
    //console.log("relMnemonic : "+relMnemonic);
    let intValue = period.toUpperCase().replace(relMnemonic,"");
    //console.log("intValue :"+intValue);
    return ! isNaN(parseInt(intValue));
  }else{
    return false;
  }
}
let fromOADate = (oadate) => {
    // let date = new Date(((oadate - 25569) * 86400000));
    // let tz = date.getTimezoneOffset();
    // return new Date(((oadate - 25569 + (tz / (60 * 24))) * 86400000));

    //todo: matching using epoch , trying to integerate momentjs instead
    return new Date((oadate - (25567+2 )) * 86400 * 1000);
  }

/**
 * Gets values from Research Management System
 * @CustomFunction
 * @param security A valid security identifier such as ticker or ticker/exchange combination, CUSIP or ISIN. For example, IBM US Equity, or 009612181.
 * @param field One or more recognized data points representing data user would like to download to Excel.
 * @param [period] The period can be of two types
    1.	Absolute reference
    FY, Q1, Q2, Q3, Q4, H1, H2
    2.	Relative reference
    LTM, YTD, LTM-1…. LTM-12,
    Q, Q-1…. Q-12,
    FY, FY-1……FY-12
    LastAvailableQ, LastAvailableQ-1……LastAvailableQ-12

 * @param [Year] The financial year for which user want to extract data. If left blank will default to the current year. Year and AsOfDate parameters can be used interchangeably but not together.
 * @param [AsOfDate] As Of Date parameter
 * @returns {any} Values received from Research Management System
 */
//@ts-ignore
function CFRMS(security: string, field: string, period: string, Year: string, AsOfDate: any) {


  console.log(`Params : security : ${security} field : ${field} period : ${period} Year : ${Year}  AsOfDate : ${AsOfDate} `);

  if (security === null || field === null) {
    console.log('Invalid number of parameters passed to function');
    throw 'Invalid number of parameters passed to function';
  }

  if (Year != null && AsOfDate != null) {
    console.log('Either year or AsOfDate can be used to invoke the formula, not both');
    console.log('Year' + (Year === ''));
    console.log('AsOfDate' + AsOfDate);

    //throw 'Either year or AsOfDate can be used to invoke the formula, not both';
    return "#N/A Either year or AsOfDate can be used get mnemonic value";
  }



    if (period != null && !isValidPeriod(period)) {
      return "#N/A Invalid period requested";
    }

    console.log("typeof AsOfDate" + typeof AsOfDate);
    if(typeof AsOfDate == "number"){
      let date = fromOADate(AsOfDate);
      AsOfDate = `${date.getMonth()+1}/${date.getDate()}/${date.getFullYear()}`

    }else if(typeof AsOfDate == "string"){
      // assuming the date format for AsOfDate is "mm/dd/yyyy"
      // since MM/DD/YYYY format is default for js it would get parsed as proper date
      let date = new Date(AsOfDate);
      AsOfDate = `${date.getMonth()+1}/${date.getDate()}/${date.getFullYear()}`

      if(isNaN(date.getDate()) || isNaN(date.getMonth()) || isNaN(date.getFullYear()) ){
        return "N/A Invalid Date"
      }

    }


    console.log(" Parsed AsOfDate :"+AsOfDate )

    return _pushOperation(
      "CFRMS",
      [security, field, period, Year, AsOfDate]
    );
  }

  /**
   * Defines the implementation of the custom functions
   * for the function id defined in the metadata file (functions.json).
   */
  CustomFunctions.associate("CFRMS", CFRMS);

  ///////////////////////////////////////

  // Next batch
  interface IBatchEntry {
    operation: string;
    args: any[];
    resolve: (data: any) => void;
    reject: (error: Error) => void;
  }

  interface IServerResponse {
    Result?: any;
    Error?: string;
    ResultType?: string;
  }

  const _batch: IBatchEntry[] = [];
  let _isBatchedRequestScheduled = false;
  let _maxConcurrentHttpRequests = 30;
  let _httpRequestCount = 0;
  let _httpRequestRetryCount = 20;
  // This function encloses your custom functions as individual entries,
  // which have some additional properties so you can keep track of whether or not
  // a request has been resolved or rejected.
  function _pushOperation(op: string, args: any[]) {
    // Create an entry for your custom function.
    const invocationEntry: IBatchEntry = {
      operation: op, // e.g. sum
      args: args,
      resolve: undefined,
      reject: undefined,
    };

    // Create a unique promise for this invocation,
    // and save its resolve and reject functions into the invocation entry.
    const promise = new Promise((resolve, reject) => {
      invocationEntry.resolve = resolve;
      invocationEntry.reject = reject;
    });

    // Push the invocation entry into the next batch.
    _batch.push(invocationEntry);

    // If a remote request hasn't been scheduled yet,
    // schedule it after a certain timeout, e.g. 100 ms.
    if (!_isBatchedRequestScheduled) {
      _isBatchedRequestScheduled = true;
      setTimeout(_scheduleRemoteRequest, 100);
    }

    // Return the promise for this invocation.
    return promise;
  }

  function _scheduleRemoteRequest(){
    while(_batch.length > 0){
      if(_batch.length > 0 && (_httpRequestCount <=_maxConcurrentHttpRequests)){
        const batchCopy = _batch.splice(0,1000);
        _makeRemoteRequest(batchCopy);
      }
    }
    _isBatchedRequestScheduled = false;
  }

  // This is a private helper function, used only within your custom function add-in.
  // You wouldn't call _makeRemoteRequest in Excel, for example.
  // This function makes a request for remote processing of the whole batch,
  // and matches the response batch to the request batch.
  function _makeRemoteRequest(batchCopy) {
    // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
    // Note the use of "splice" rather than "slice", which will modify the original _batch array
    // to empty it out.
    // in optimistic scenarios, we can send the whole batch of accumulated requests
    // but to handle this scenario to reduce network load, we are batching the request in smaller chunks
    //const batchCopy = _batch.splice(0, _batch.length);
    // todo: remove this hard-coding to get it from server instead


    // Build a simpler request batch that only contains the arguments for each invocation.
    const requestBatch = batchCopy.map((item, index) => {
      return { operation: item.operation, args: item.args, index: index };
    });

    let authToken = "";
    //@ts-ignore
    OfficeRuntime.storage.getItem('token').then((authentication_token) => {
      authToken = authentication_token;
      _httpRequestCount++;
       // Make the remote request.
    _fetchFromRemoteService(requestBatch, authToken)
    .then((responseBatch) => {
      _httpRequestCount--;
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      console.log(responseBatch);


      responseBatch.forEach((response, index) => {
        if (response.Error) {
          //batchCopy[index].reject(new Error(response.error));
          // todo: see if we can re-trigger the formula calculation in this case
          batchCopy[index].resolve(String(response.Error));
        } else {
          if(response.ResultType === 'number')
            batchCopy[index].resolve(Number(response.Result));
          else
            batchCopy[index].resolve(String(response.Result));
        }
      });
    });

    });


  }

  function fetchWithRetryDelayMechanism(url,options) {
    return new Promise((resolve, reject) => {
      let attempts = 1;
      const fetch_retry = (url,options, n) => {
        return fetch(url,options).then(resolve).catch(function (error) {
        if (n === 1) reject(error)
        else
        setTimeout(() => {
          attempts ++
          fetch_retry(url,options, n - 1);
        }, attempts * 3000)
      });
    }
      return fetch_retry(url,options, _httpRequestRetryCount);
    });
  }

  async function _fetchFromRemoteService(
    requestBatch: Array<{ operation: string, args: any[] }>
    , authToken: string
  ): Promise<IServerResponse[]> {

    //@ts-ignore
    let url = `${API_URL}/ExcelBatchFormulaHandler`
    //let formulaValues =JSON.stringify(requestBatch);

    //@ts-ignore
    //OfficeRuntime.storage.getItem('token').then((authentication_token) => {
    let formulaValues = { formulaValues: requestBatch, authToken: authToken };
    let serverRequest = JSON.stringify(formulaValues);

    return await fetchWithRetryDelayMechanism(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain'
      },
      body: serverRequest
    }).then(function (response:Response) {
        return response.json();
      }
      ).then(function (serverResponse) {

        return serverResponse.map((response): IServerResponse => {
          if (response.error) {
            return {
              Error: response.error,
              ResultType : "text"
            };
          } else {
            return {
              Result: response.result,
              ResultType : response.resultType
             };
          }
        });
      })
      .catch(function (error) {
        console.log('error', error.message);
        // send the error back to server to log user level exceptions
        //@ts-ignore
        fetch(`${API_URL}/excelerror`, {
          method: 'POST',
          headers: {
            'Content-Type': 'text/plain'
          },
          body: JSON.stringify(error.message)
        });
        return requestBatch.map((request): IServerResponse => {
          return {
            Error: `#N/A Operation failed` // - ${error.message}`
          };
        });
      });

  }

  function pause(ms: number) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
