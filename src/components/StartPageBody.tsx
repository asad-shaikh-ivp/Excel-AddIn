import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import HeroList, { HeroListItem } from './HeroList';

export interface StartPageBodyProps {
    listItems: HeroListItem[];
    login: () => {};
}

export default class StartPageBody extends React.Component<StartPageBodyProps> {
    render() {
        const { listItems, login } = this.props;

        return (
            <div className='ms-welcome'>

                <div className='ms-welcome__main'>
                    <HeroList message='Sign in to proceed further.' items={listItems}>
                    </HeroList>
                    <Button className='ms-welcome__actionPrimary' buttonType={ButtonType.hero} onClick={login}>Sign In</Button>
                </div>
            </div>
        );
    }
}
