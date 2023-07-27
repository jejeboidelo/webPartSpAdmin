
import { Button } from '@fluentui/react-northstar';
import { SPFI } from '@pnp/sp';
import * as React from 'react';
import { getSP } from '../../../PnPJsConfig';
import AdminServices from '../../services/AdminServices';


interface ISiteLockProps {
    disabled: boolean,
    site: any
  }
  
interface ISiteLockState {
    disabled: boolean,
    etatActuel: string,
    loading: boolean,
}
  
export default class SiteLockNoAccess extends React.Component<ISiteLockProps, ISiteLockState> {
  private _sp: SPFI;
  
    public constructor(props: ISiteLockProps) {
        super(props);
        
        this.state = {
            disabled: props.disabled,
            etatActuel: props.site.LockState,
            loading: false
        };

        this._sp = getSP();
        this.locking = this.locking.bind(this);
      }

      public render(): React.ReactElement<ISiteLockProps> {
        return(
          <Button
          size='small'
          onClick={this.locking}
          loading={this.state.loading}
          disabled={this.state.disabled || this.state.etatActuel=="NoAccess"}
          >NoAccess</Button>  
        )
      }

      private async locking() {
        this.setState({
          loading: true
        })

        let id = null;

        try {
          id = await AdminServices.getId(this.props.site.Title)
        } catch (e) {
          alert("Erreur lors de la récupération de l'Id du site")
        }
        
        if(id!= null){
          try{
            await this._sp.web.lists.getByTitle("IDSite").items.add({
              Title: this.props.site.Title,
              IDsite: id,
              GroupID: this.props.site.GroupId,
              Url: this.props.site.Url,
              Template: this.props.site.Template,
              LockState: this.props.site.LockState
            });
          }
          catch (e){
            id = null
            alert("Erreur lors de l'inscription dans la liste")
          }
        }
        if(id != null){
          try{
            console.log("##"+this.props.site.Title+"###");
            await AdminServices.UpdateSiteProperties(this.props.site.Title, { "LockState": "NoAccess"});
          }catch (e){
            alert("Erreur lors du passage du site en innaccessible")
          }}  
        
        this.setState({
          etatActuel:"NoAccess",
          loading: false
        })
      }
}