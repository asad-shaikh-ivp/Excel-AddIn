import * as React from "react";
import { Component } from "react";
//import { Button, ButtonType } from "office-ui-fabric-react";
import Button from '@material-ui/core/Button';
import HeroList, { HeroListItem } from "./HeroList";
import { withStyles } from "@material-ui/core/styles";
import Table from "@material-ui/core/Table";
import TableBody from "@material-ui/core/TableBody";
import TableCell from "@material-ui/core/TableCell";
import TableContainer from "@material-ui/core/TableContainer";
import TableHead from "@material-ui/core/TableHead";
import TableRow from "@material-ui/core/TableRow";
import Paper from "@material-ui/core/Paper";
import Typography from '@material-ui/core/Typography';


const useStyles = (theme) => ({

    root: {
      '& > *': {
        margin: theme.spacing(1),
        justifyContent: 'center'
      },
    }, 
    typography: {
      
      fontSize: 10,
    },

  table: {},

});

export interface MnemonicListBodyProps {
  setParentState: () => {};
}

class MnemonicListBody extends React.Component<
  MnemonicListBodyProps
> {
  state = {
    hasErrors: false,
    mnemonic_name: [],
  };

  GetSortOrder = (prop) => {
    return function (a, b) {
      if (a[prop] > b[prop]) {
        return 1;
      } else if (a[prop] < b[prop]) {
        return -1;
      }
      return 0;
    };
  };

  componentDidMount() {

    OfficeRuntime.storage.getItem('token').then((authentication_token) => {
      const requestHeaders: HeadersInit = new Headers();

      requestHeaders.append('Authorization', 'Bearer ' + authentication_token);
      requestHeaders.append('Content-Type', 'application/json');

      //@ts-ignore
      let url = `${API_URL}/api/GetMnemonicList`;

      fetch(url, { method: 'GET', 'headers': requestHeaders })
      .then((res) => res.json())
      .then((res) => {
        for (let i = 0; i < res.mnemonicList.length; i++) {
          if (res.mnemonicList[i].mnemonicCalcType == "NOOP") {
            res.mnemonicList[i].mnemonicCalcType = "-";
          }
        }

        res.mnemonicList.sort(this.GetSortOrder("mnemonicName"));
        this.setState({
          mnemonic_name: res.mnemonicList,
        });
      })
      .catch(() => this.setState({ hasErrors: true }));
    });
  }

  render() {

    //@ts-ignore
    const {classes} = this.props;
    //@ts-ignore
    const { setParentState } = this.props;



    return (
      <div className={classes.root}>
        <div className={classes.root}>
            <div >
          {/* <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            onClick={setParentState}
          >
            Back
          </Button> */}

          <Button variant="contained" onClick={setParentState} disableElevation>Back</Button>
          </div>
          <br />
            <TableContainer component={Paper}>
              <Table
                className={classes.table}
                size="small"
                aria-label="a dense table"
              >
                <TableHead>
                  <TableRow>
                    <TableCell style = {{width: '56%', fontWeight: 550, }}>Mnemonic Name </TableCell>
                    <TableCell style = {{width: '22%', fontWeight: 550,}}>Calc Type</TableCell>
                    <TableCell style = {{width: '22%', fontWeight: 550,}}>Data Type</TableCell>
                  </TableRow>
                </TableHead>
              <TableBody>
                {

                  //@ts-ignore
                  this.state.mnemonic_name.map((mnemonic) => (
                    <TableRow key={mnemonic.mnemonicName}>
                      <TableCell component="td" scope="row" style = {{fontSize: 11.5}}>
                        {mnemonic.mnemonicName}
                      </TableCell>
                      <TableCell style = {{fontSize: 11.5}}>
                        {mnemonic.mnemonicCalcType}
                      </TableCell>
                      <TableCell style = {{fontSize: 11.5}}>
                        {mnemonic.mnemonicTypeName}
                      </TableCell >
                    </TableRow>
                  ))
                }
              </TableBody>
              </Table>
            </TableContainer>


          {/* <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            onClick={this.props.setParentState}
          >
            Back
          </Button> */}
          <Button variant="contained" onClick={setParentState} disableElevation>Back</Button>
        </div>
      </div>
    );
  }
}

export default withStyles(useStyles)(MnemonicListBody)