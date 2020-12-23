import * as React from 'react';
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
export interface HeaderProps {
    title: string;
    logo: string;
    message: string;
    userName?: string;
}

export default class Header extends React.Component<HeaderProps> {
    render() {
        const { title, logo, message, userName } = this.props;

        return (
            <section className='ms-welcome__header ms-bgColor-themeSecondary ms-u-fadeIn500'>

                 <img width='' height='30' src={logo} alt={title} title={title} />
                 <h1 className='ms-fontSize-xl ms-fontWeight-light ms-fontColor-white'>Research Management System</h1>
                 <h1 className='ms-fontSize-xl ms-fontWeight-light ms-fontColor-white'>{message}</h1>
                 { userName?(
                    <h1 className='ms-fontSize-xl ms-fontWeight-light ms-fontColor-white'>Signed in as {userName}</h1>
                    ): (<h1></h1>)
                }
             </section>
        );
    }
}
