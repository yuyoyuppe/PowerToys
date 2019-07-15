import React from 'react';
import { Stack, Text, Toggle } from 'office-ui-fabric-react';

export class GeneralSettings extends React.Component <any, any> {
  references: any;
  startup_reference: any;
  constructor(props: any) {
    super(props);
    this.references={};
    this.startup_reference=null;
    this.state = {
      settings_key: props.settings_key,
      settings: props.settings,
    }
  }
/*
  static getDerivedStateFromProps(props: any, state: any) {
    if (state.settings!=props.settings) {
      alert('b');
      return { settings: props.settings };
    }
    return null;
  }
*/
  public get_data(): any {
    let enabled : any = {};
    Object.keys(this.references).forEach(key => {
      enabled[key]=this.references[key].checked;
    });
    let result : any = {};
    result[this.state.settings_key]= {
      startup: this.startup_reference.checked,
      enabled: enabled
    };
    return result;
  }

  public render(): JSX.Element {
    let power_toys_enabled = this.state.settings.enabled;
    return (
      <Stack tokens={{childrenGap:30}}>
        <Text variant='xLarge'>Enabled PowerToys</Text>
        { Object.keys(power_toys_enabled).map(
          (key) => {
            let enabled_value=power_toys_enabled[key];
            return <Toggle
              key={key}
              defaultChecked={enabled_value}
              label={key}
              onText="Enabled"
              offText="Disabled"
              componentRef={(input) => {this.references[key]=input;}}
            /> ;
          })
        }
        <Text variant='xLarge'>General</Text>
        <Toggle
          defaultChecked={this.state.startup}
          label="Start at login"
          onText="Enabled"
          offText="Disabled"
          componentRef= {(input) => {this.startup_reference=input;}}
        />
      </Stack>
    )
  }
}
