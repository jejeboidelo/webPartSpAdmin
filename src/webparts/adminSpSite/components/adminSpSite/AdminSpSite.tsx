import * as React from 'react';

import { IAdminSpSiteProps } from './IAdminSpSiteProps';
import Sites from '../Sites';
import HubSites from '../HubSites';
import { Button, Divider, Provider, RadioGroup, RadioGroupItemProps, ShorthandCollection, teamsTheme } from '@fluentui/react-northstar';


enum Menu {
  Site,
  HubSites
}

interface IAdminState {
  selectedMenu: Menu;
}


export default class AdminSpSite extends React.Component<IAdminSpSiteProps, IAdminState> {

  private readonly _menu: ShorthandCollection<RadioGroupItemProps> = [
    {
      key: "Sites",
      styles: { width: '100px' },
      checkedIndicator: <Button primary content="Sites" />,
      indicator: <Button secondary content="Sites" />,
      value: Menu.Site
    },
    {
      key: "HubSite",
      styles: { width: '100px' },
      checkedIndicator: <Button primary content="HubSite" />,
      indicator: <Button secondary content="HubSite" />,
      value: Menu.HubSites
    }]


  public constructor(props: IAdminSpSiteProps){
    super(props);
    this.state = {
      selectedMenu: Menu.Site
    }
  }
  public render(): React.ReactElement<IAdminSpSiteProps> {
   

    return (
      <Provider theme={teamsTheme}>
        <RadioGroup
          checkedValue={this.state.selectedMenu}
          defaultCheckedValue={Menu.Site}
          onCheckedValueChange={(_ev, data)=>{this.setState({ selectedMenu: data.value as Menu });}}
          styles={{ padding: '40px' }}
          items={this._menu} />
          
        <Divider styles={{ paddingBottom: "15px" }} />
        {this.state.selectedMenu === Menu.Site &&
        <Sites spHttpClient={this.props.context.spHttpClient}/>}
        {this.state.selectedMenu === Menu.HubSites &&
        <HubSites/>}
      </Provider>
      
    );
  } 
}
