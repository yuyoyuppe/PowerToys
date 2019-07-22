import React from 'react';
import {BaseSettingsControl} from './BaseSettingsControl';
import {Label, Stack, PrimaryButton, TextField} from 'office-ui-fabric-react';

export class CustomActionSettingsControl extends BaseSettingsControl {
  colorpickerref:any = null;

  constructor(props:any) {
    super(props);
    this.colorpickerref = null;
    this.state={
      property_values: props.setting,
      call_action_callback: props.action_callback,
      name: props.action_name
    }
  }

  componentWillReceiveProps(props: any) {
    this.setState({
      property_values: props.setting,
      name:props.action_name
    });
  }

  public get_value() : any {
    return {value: this.state.property_values.value};
  }

  public render(): JSX.Element {
    return (
      <Stack>
        <Label>{this.state.property_values.display_name}</Label>
        {
          this.state.property_values.value ?
          <TextField
            styles = {{
              root: {
                paddingBottom: '5px',
              },
              fieldGroup: {
                minHeight: '0px',
                borderColor: '#d0d0d0',
                selectors: {
                  '&:hover': {
                    borderColor: '#d0d0d0'
                  }
                }
              },
              field: {
                height: '0px' // To override the initial height set by fabric for a multiline textfield.
              }
            }}
            multiline={true}
            borderless={false}
            autoAdjustHeight={true}
            readOnly={true}
            resizable={false}
            value={this.state.property_values.value}
          /> :
          <span/>
        }
        <PrimaryButton
            styles={{
              root: {
                alignSelf: 'start'
              }
          }}
          text={this.state.property_values.button_text}
          onClick={()=>this.state.call_action_callback(this.state.name, this.state.property_values)}
        />
      </Stack>
    );
  }
}
