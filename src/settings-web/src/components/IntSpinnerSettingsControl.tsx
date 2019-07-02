import React from 'react';
import {BaseSettingsControl} from './BaseSettingsControl';
import { SpinButton } from 'office-ui-fabric-react';
import { Position } from 'office-ui-fabric-react/lib/utilities/positioning';

export class IntSpinnerSettingsControl extends BaseSettingsControl {
  spinbuttonref:any = null;

  constructor(props:any) {
    super(props);
    this.spinbuttonref = null;
    this.state={
      property_values: props.setting
    }
  }

  public get_value() : any {
    return {value: parseInt(this.spinbuttonref.value)};
  }

  public render(): JSX.Element {
    return (
      <SpinButton
        defaultValue={this.state.property_values.value}
        precision={0}
        step={1}
        label={this.state.property_values.display_name}
        labelPosition={Position.top}
        componentRef= {(input) => {this.spinbuttonref=input;}}
      />
    );
  }

}