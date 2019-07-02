import React from 'react';
import {BaseSettingsControl} from './BaseSettingsControl';
import { ColorPicker, Stack, Label } from 'office-ui-fabric-react';

export class ColorPickerSettingsControl extends BaseSettingsControl {
  colorpickerref:any = null;

  constructor(props:any) {
    super(props);
    this.colorpickerref = null;
    this.state={
      property_values: props.setting
    }
  }

  public get_value() : any {
    return {value: this.colorpickerref.color.str};
  }

  public render(): JSX.Element {
    return (
      <Stack>
        <Label>{this.state.property_values.display_name}</Label>
        <ColorPicker
          styles= {{
            panel: {padding:0}
          }}
          color={this.state.property_values.value}
          componentRef= {(input) => {this.colorpickerref=input;}}
          alphaSliderHidden = {true}
        />
      </Stack>
    );
  }

}