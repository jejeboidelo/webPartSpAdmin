import { Tooltip, WandIcon } from '@fluentui/react-northstar';
import { Button, LockIcon } from '@fluentui/react-northstar';
import * as React from 'react';
import AdminServices from '../../services/AdminServices';

interface ISiteLockProps {
    etatActuel: string,
    siteName: string,
    disabled: boolean
  }
  
interface ISiteLockState {
    etatActuel: string,
    loading: boolean,
}
  
export default class SiteLock extends React.Component<ISiteLockProps, ISiteLockState> {

    public constructor(props: ISiteLockProps) {
        super(props);
        
        this.state = {
            etatActuel: props.etatActuel,
            loading: false
        };

        this.locking = this.locking.bind(this);
      }

      public render(): React.ReactElement<ISiteLockProps> {
        return(
            <Tooltip
            trigger={
                <Button
                icon={this.state.etatActuel === "Unlock"?<LockIcon/>:<WandIcon/>}
                onClick={this.locking}
                loading={this.state.loading}
                disabled={this.props.disabled}
                iconOnly
                />}
            content= {this.state.etatActuel === "Unlock"?"Lecture Seule":"Modifiable"}
            />

            
        )
      }

      private async locking() {
        this.setState({
          loading: true
        })
        console.log("##"+this.props.siteName+"###");
        
        await AdminServices.UpdateSiteProperties(this.props.siteName, { "LockState": this.state.etatActuel === "Unlock"?"ReadOnly":"Unlock"});  
        
        this.setState({
          etatActuel:(this.state.etatActuel === "Unlock"?"ReadOnly":"Unlock"),
          loading: false
        })
      }
}