import * as React from 'react';
import { Spinner, SpinnerType } from 'office-ui-fabric-react';
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
        //primaryBackground: '#fff',
        //primaryText: '#333'
    }
});
export interface ProgressProps {
    logo: string;
    message: string;
    title: string;
}

export default class Progress extends React.Component<ProgressProps> {
    render() {
        const { logo, message, title } = this.props;

        return (
            <section className='ms-welcome__progress ms-u-fadeIn500'>
                <img width='90' height='90' src={logo} alt={title} title={title} />
                <h1 className='ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary'>{title}</h1>
                <Spinner type={SpinnerType.large} label={message} />
            </section>
        );
    }
}
