import React from 'react';
import {BaseSettingsControl} from './BaseSettingsControl';
import { Toggle } from 'office-ui-fabric-react';

export class BoolToggleSettingsControl extends BaseSettingsControl {
  toggleref:any = null;

  constructor(props:any) {
    super(props);
    this.toggleref = null;
    this.state={
      property_values: props.setting
    }
  }

  public get_value() : any {
    return {value: this.toggleref.checked};
  }

  public render(): JSX.Element {
    return (
      <Toggle
        defaultChecked={this.state.property_values.value}
        label={this.state.property_values.display_name}
        onText="Enabled"
        offText="Disabled"
        componentRef= {(input) => {this.toggleref=input;}}
      />
    );
  }

}