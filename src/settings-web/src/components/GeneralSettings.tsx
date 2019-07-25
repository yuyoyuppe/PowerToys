import React from 'react';
import { Stack, Text, DefaultButton, Label} from 'office-ui-fabric-react';
import {BoolToggleSettingsControl} from './BoolToggleSettingsControl'
import { Separator } from 'office-ui-fabric-react/lib/Separator';

export class GeneralSettings extends React.Component <any, any> {
  references: any = {};
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

  componentWillReceiveProps(props: any) {
    this.setState({ settings: props.settings })
  }

  public get_data(): any {
    let enabled : any = {};
    Object.keys(this.references).forEach(key => {
      enabled[key]=this.references[key].get_value().value;
    });
    let result : any = {};
    result[this.state.settings_key]= {
      startup: this.startup_reference.get_value().value,
      enabled: enabled
    };
    return result;
  }

  public render(): JSX.Element {
    let power_toys_enabled = this.state.settings.enabled;
    return (
      <Stack tokens={{childrenGap:20}}>
        <Text variant='xLarge'>Available PowerToys</Text>
        { Object.keys(power_toys_enabled).map(
          (key) => {
            let enabled_value=power_toys_enabled[key];
            return <BoolToggleSettingsControl
              setting={{display_name: key, value: enabled_value}}
              key={key}
              ref={(input) => {this.references[key]=input;}}
            />;
          })
        }
        <Separator />
        <Text variant='xLarge'>General</Text>
        <BoolToggleSettingsControl
          setting={{display_name: 'Start at login', value: this.state.settings.startup}}
          ref={(input) => {this.startup_reference=input;}}
          />
        <Stack>
        <Label>Version 0.9.0</Label>
          <DefaultButton
            styles={{
                root: {
                  backgroundColor: "#FFFFFF",
                  alignSelf: "start"
                }
            }}
            href='https://github.com/microsoft/PowerToys/releases'
            target='_blank'
          >Check for updates</DefaultButton>
        </Stack>
        {/* An empty span to always give 30px padding in Edge. */}
        <span/>
      </Stack>
    )
  }
}
