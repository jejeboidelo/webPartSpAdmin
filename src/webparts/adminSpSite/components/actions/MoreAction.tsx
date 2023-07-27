/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @microsoft/spfx/no-async-await */
/* eslint-disable react/jsx-no-bind */
import * as React from 'react';
import {
  Header,
  Label,
  Checkbox,
  Dropdown,
  Button,
  Flex,
  Loader,
  Divider,
} from "@fluentui/react-northstar";
import { Caching } from "@pnp/queryable";
import AdminServices from '../../services/AdminServices';
import { FlowsPolicy, ITenantSitePropertiesInfo, SharingCapabilities } from '@pnp/sp-admin';
import { Panel } from 'office-ui-fabric-react';
import { getSP } from '../../../PnPJsConfig';
import { SPHttpClient } from '@microsoft/sp-http';
// import { Web } from '@pnp/sp/webs';
import "@pnp/sp/hubsites/site";


export interface IMoreActionProps {
  spHttpCtx: SPHttpClient;
  isPanelOpen: boolean;
  site: Partial<ITenantSitePropertiesInfo>;
  isHomeSite: boolean;
  isCommSite: boolean;
  closePanel: () => void;
  refreshHomeSite: (newSiteUrl: string) => void;
}

interface IMoreActionState {
  isLoading: boolean;
  isNewHomeSite: boolean;
  commentDisabled: boolean;
  disableFlow: number;
  site: any;
}

export default class MoreAction extends React.Component<IMoreActionProps, IMoreActionState> {
  private readonly flowPolicies = (Object.keys(FlowsPolicy).filter(value => isNaN(Number(value)) === false).map((key: any) => FlowsPolicy[key]));
  private readonly sharingCapabilities = (Object.keys(SharingCapabilities).filter(value => isNaN(Number(value)) === false).map((key: any) => SharingCapabilities[key]));

  public constructor(props: IMoreActionProps) {
    super(props);
    
    this.state = {
      isLoading: false,
      isNewHomeSite: false,
      commentDisabled: this.props.site?.CommentsOnSitePagesDisabled,
      disableFlow: this.props.site?.DisableFlows,
      site: null
    };
  }

  public render(): React.ReactElement<IMoreActionProps> {
    return (
      <Panel
        styles={{ main: { width: "500px !important" } }} // TODO: get styles from Northstar theme to apply
        closeButtonAriaLabel="Close"
        isOpen={this.props.isPanelOpen}
        onOpened={()=> {this.setState({commentDisabled: this.props.site?.CommentsOnSitePagesDisabled, disableFlow: this.props.site?.DisableFlows, site: this.props.site});
                        console.log(this.state.disableFlow);}}
        onDismiss={(_ev) => { this.setState({ isNewHomeSite: false }); this.props.closePanel() }}>

        <Header>Site properties</Header>
        <p>{this.state.commentDisabled}</p>
        <b>Site: </b><Label circular color="brand">{this.props.site?.Title}</Label>
        <Divider styles={{ paddingTop: "15px", paddingBottom: "15px" }} />
        {this.state.isLoading
          ?
          <Loader label="Applying changes..." />
          :
          <Flex gap="gap.small" column>
            {this.props.isCommSite &&
              <Checkbox labelPosition='start' styles={{ maxWidth: "300px", paddingLeft: "0px" }} toggle label="Comments on site pages disabled" checked={this.state.commentDisabled} onChange={(_ev, data) => {this._updateSiteProperty({ "CommentsOnSitePagesDisabled": data.checked }); this.setState({commentDisabled: data.checked})}} />
            }
            {!this.props.isHomeSite && this.props.isCommSite &&
              <Checkbox labelPosition='start' disabled={this.state.isNewHomeSite} styles={{ maxWidth: "300px", paddingLeft: "0px" }} label="Set this site as Home Site" checked={this.state.isNewHomeSite} onChange={this._setHomeSite} />
            }
            <span>
              Flow policy:{' '}
              <Dropdown inline items={this.flowPolicies} defaultValue={FlowsPolicy[this.state.disableFlow]} onChange={(_ev, data) => {this._updateSiteProperty({ "DisableFlows": data.highlightedIndex });this.setState({disableFlow: data.highlightedIndex})}} />
            </span>
            <span>
              Sharing Capabilities:{' '}
              <Dropdown inline items={this.sharingCapabilities} defaultValue={SharingCapabilities[this.props.site?.SharingCapability]} onChange={(_ev, data) => {this._updateSiteProperty({ "SharingCapability": data.highlightedIndex })}} />
            </span>
            <span>ajouter en tant quie site hub :
              <Button onClick={this.newhubSite}></Button>
            </span>
          </Flex>
        }
        <Flex gap='gap.small' styles={{ position: "absolute", bottom: "0", paddingBottom: "15px" }}>
          <Button content="Close" type='button' secondary onClick={(_ev, _data) => this.props.closePanel()} disabled={this.state.isLoading} />
        </Flex>
      </Panel>
    );
  }

  private _updateSiteProperty = async (updatedProperty: Record<string,any>): Promise<void> => {
    this.setState({
      isLoading: true,
    });

    try {

      await AdminServices.UpdateSiteProperties(this.props.site.Title, updatedProperty);

      await AdminServices.ajoutItemList(
        {Title: this.props.site.Title,
        GroupID: this.props.site.GroupId,
        Url: this.props.site.Url,
        Template: this.props.site.Template,
        ...updatedProperty
        }
      );

    } catch (error) {
      console.log(error);
    }
    finally {
      this.setState({
        isLoading: false,
      });
    }
  }

  private _setHomeSite = async (): Promise<void> => {

    try {
      this.setState({
        isLoading: true,
      });

      const contextUrl: string = (await getSP().using(Caching()).site()).Url;

      await (await this.props.spHttpCtx.post(
        contextUrl + "/_api/SPHSite/SetSPHSite",
        SPHttpClient.configurations.v1,
        {
          body: JSON.stringify({
            siteUrl: this.props.site.Url
          }),
        })).json();

      this.setState({
        isNewHomeSite: true
      });

      this.props.refreshHomeSite(this.props.site.Url);
    } catch (error) {
      console.error(error);
    }
    finally {
      this.setState({
        isLoading: false,
      });
    }
  }

  public async newhubSite(){
    // const _sp = getSP()

    // const w = await _sp.site.openWebById(this.props.site?.Url);
    // w.
    // const _web = Web(this.props.site?.Url)
    // _web.reg
    // console.log(this.props.site);

    // console.log(this.state.site.Url);
    // console.log(this.state.site?.Url+"/_api/site/RegisterHubSite");

    await fetch("https://yzjlx.sharepoint.com/sites/testpourhomesite/_api/site/RegisterHubSite", {
      method: 'POST',
      body: null,
      headers: {
        "Accept": "application/json;odata=verbose", 
        "Content-Type": "application/json;odata=verbose",
        "X-HTTP-Method": "MERGE"
      }
    })
      .then(res => console.log(res))
      .catch(err => console.log(err))
  }
}