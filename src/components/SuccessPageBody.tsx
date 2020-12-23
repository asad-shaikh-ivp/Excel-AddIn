import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';

export interface SuccessPageBodyProps {
    logout: () => {};
    uploadModel : () => {};
    deleteModel : () => {}; 
    mnemonicList : () => {};
}


export default class SuccessPageBody extends React.Component<SuccessPageBodyProps> {

  

    render() {
        //@ts-ignore
        const {  logout, uploadModel, deleteModel, mnemonicList } = this.props;


    //     let key = 'token';
    //   let tokenSendStatus = 'tokenSendStatus :';
    //   let tokenValue = localStorage.getItem('msal.idtoken');
    //   console.log(tokenValue);

    //   OfficeRuntime.storage.setItem(key, tokenValue).then((result) => {
    //       tokenSendStatus = 'Success: ' + key + result;
    //       console.log(tokenSendStatus);
    //   }, (error) => {
    //         tokenSendStatus = 'failure : ' + key + error;
    //       console.log(tokenSendStatus);
    //   });

        return (
            <div className='ms-welcome__main'>
                <h2 className='ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20'></h2>
                <Button  className='ms-welcome__action' buttonType={ButtonType.hero}  onClick={mnemonicList}>View Mnemonics</Button>
                <Button className='ms-welcome__actionPrimary' buttonType={ButtonType.hero}  onClick={uploadModel}>Upload</Button>
                <Button className='ms-welcome__action' buttonType={ButtonType.hero} onClick={deleteModel}>Delete</Button> 
                {/* style={{backgroundColor:'red', color:'white'}}  */}
                <Button className='ms-welcome__action' buttonType={ButtonType.hero}  onClick={logout}>Sign out</Button>
            </div>
        );
    }
}
