import React from 'react';
import {BaseSettingsControl} from './BaseSettingsControl';
import { TextField } from 'office-ui-fabric-react';

export class StringTextSettingsControl extends BaseSettingsControl {
  textref:any = null;

  constructor(props:any) {
    super(props);
    this.textref = null;
    this.state={
      property_values: props.setting
    }
  }

  public get_value() : any {
    return {value: this.textref.value};
  }

  public render(): JSX.Element {
    return (
      <TextField
        defaultValue={this.state.property_values.value}
        label={this.state.property_values.display_name}        
        componentRef= {(input) => {this.textref=input;}}
      />
    );
  }

}