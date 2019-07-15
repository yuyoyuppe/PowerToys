import React from 'react';
import { Stack } from 'office-ui-fabric-react';
import {BoolToggleSettingsControl} from './BoolToggleSettingsControl';
import {StringTextSettingsControl} from './StringTextSettingsControl';
import {IntSpinnerSettingsControl} from './IntSpinnerSettingsControl';
import {ColorPickerSettingsControl} from './ColorPickerSettingsControl';

export class CustomSettingsScreen extends React.Component <any, any> {
  references: any;

  constructor(props: any) {
    super(props);
    this.references={};
    this.state = {
      settings_key: props.settings_key,
      powertoy: props.powertoy,
    }
  }
  componentWillReceiveProps(props: any) {
    this.setState({ powertoy: props.powertoy })
  }

  public get_data(): any {
    let properties : any = {};
    Object.keys(this.references).forEach(key => {
      properties[key]= this.references[key].get_value();
    });
    let result : any = {};
    result[this.state.settings_key] = {
      name: this.state.powertoy.name,
      properties:properties
    };
    return {powertoys: result};
  }

  public render(): JSX.Element {
    let power_toys_properties = this.state.powertoy.properties;
    return (
      <Stack tokens={{childrenGap:30}}>
        { Object.keys(power_toys_properties).map(
          (key) => {
            switch(power_toys_properties[key].editor_type) {
              case 'bool_toggle':
                return <BoolToggleSettingsControl
                  setting={power_toys_properties[key]}
                  key={key}
                  ref={(input) => {this.references[key]=input;}}
                  />;
              case 'string_text':
                return <StringTextSettingsControl
                  setting = {power_toys_properties[key]}
                  key={key}
                  ref={(input) => {this.references[key]=input;}}
                  />;
              case 'int_spinner':
                return <IntSpinnerSettingsControl
                  setting = {power_toys_properties[key]}
                  key={key}
                  ref={(input) => {this.references[key]=input;}}
                  />;
              case 'color_picker':
                return <ColorPickerSettingsControl
                  setting = {power_toys_properties[key]}
                  key={key}
                  ref={(input) => {this.references[key]=input;}}
                  />;
              default:
                return null;
            }
          })
        }
      </Stack>
    )
  }
}
