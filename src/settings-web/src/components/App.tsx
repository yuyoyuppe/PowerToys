import React from 'react';
import {Stack, Text, Nav, CommandButton, DefaultButton, PrimaryButton, IconButton, ScrollablePane, INavLink} from 'office-ui-fabric-react';
import {GeneralSettings} from './GeneralSettings';
import {CustomSettingsScreen} from './CustomSettingsScreen';
import '../css/layout.css';
import '../icons/css/fabric-icons-inline.css';
import {initializeIcons} from '../icons/src';
initializeIcons('src/icons/fonts/');

export class App extends React.Component <any, any> {
  settingsscreenref:any;
  constructor(props: any) {
    super(props);
    this.settingsscreenref = null;
    this.state = {
      selectedmenu : 'general',
      settings: {
        general: {
          startup: true,
          enabled: {
            'Move To New Desktop':true,
            'Shortcut Guide':false,
          }
        },
        powertoys: {
          'Move To New Desktop' : {
            name: 'Move To New Desktop',
            description: 'Adds popup that Maximizes a Window to a new Desktop.',
            properties: {
              'close desktop on restore' : {
                display_name: 'Remove a virtual desktop when the last window is restored',
                editor_type: 'bool_toggle',
                value: true
              },
              'close desktop on window close' : {
                display_name: 'Remove a virtual desktop when the last window is closed',
                editor_type: 'bool_toggle',
                value: false
              },
            },
          },
          'Shortcut Guide': {
            name: 'Shortcut Guide',
            description: 'Shows a help overlay with Windows shortcuts when the Windows key is pressed.',
            properties: {
              'press time' : {
                display_name: 'How long to press the Windows key before showing the Shortcut Guide (ms)',
                editor_type: 'int_spinner',
                value: 300
              },
            }
          },
          'Example PowerToy': {
            name: 'Example PowerToy',
            description: 'Shows the different controls for the settings.',
            properties: {
              'test bool_toggle': {
                display_name: 'This is what a bool_toggle looks like',
                editor_type: 'bool_toggle',
                value: false
              },
              'test int_spinner': {
                display_name: 'This is what a int_spinner looks like',
                editor_type: 'int_spinner',
                value: 10
              },
              'test string_text': {
                display_name: 'This is what a string_text looks like',
                editor_type: 'string_text',
                value: 'A sample string value'
              },
              'test color_picker': {
                display_name: 'This is what a color_picker looks like',
                editor_type: 'color_picker',
                value: '#0450fd'
              },
            }
          }
        }
      }
    }
  }

  public render(): JSX.Element {
    const powertoys_dict = this.state.settings.powertoys;
    let powertoys_links = [];
    for(let powertoy_key in powertoys_dict) {
      if(powertoys_dict.hasOwnProperty(powertoy_key)) {
        powertoys_links.push({
          name: powertoys_dict[powertoy_key].name,
          key: powertoy_key,
          url:'',
          icon:'CircleRing'
        });
      }
    }

    const saveClicked = (): void => {
      if (typeof (window.external) !== 'undefined' && ('notify' in window.external)) {
        (window.external as any).notify(JSON.stringify(this.settingsscreenref.get_data()));
      } else {
        alert(JSON.stringify(this.settingsscreenref.get_data()));
      }
    };
    const discardChanges = (): void => {
      this.settingsscreenref.forceUpdate();
    }

    return (
      <div className='body'>
        <div className='sidebar'>
          <CommandButton
            iconProps={{iconName: 'GlobalNavButton'}}
            text='PowerToys'
            styles={{
              textContainer: {fontSize:18},
            }}
            />
          <Nav
            selectedKey= {this.state.selectedmenu}
            onLinkClick = {
              (ev?: React.MouseEvent<HTMLElement,MouseEvent>, item?: INavLink) => {
                this.setState({selectedmenu : ((item && item.key)||null) });
              }
            }
            styles = {{
              compositeLink: {
                backgroundColor : '#f3f2f1',
                color: '#323130',
                selectors: {
                  '&.is-selected button' : {
                    backgroundColor: '#e1dfdd',
                    color: '#201F1E',
                    fontWeight: 'bold'
                  },
                  '&:hover button.ms-Nav-link' : {
                    backgroundColor: '#e1dfdd',
                    color: '#323130'
                  },
                },
              },
            }}
            groups = {[
              {
                links: powertoys_links,
              },
              {
                links: [
                  { name: 'General', key:'general', url:'', icon: 'Settings' },
                ],
              }
            ]}
          />
        </div>
        <div className='editorzone'>
          <div className='editorhead'>
            <div className='editortitle'>
              <Text
                variant='xxLarge'
                styles= {{ root: { display:'block', whiteSpace:'no-wrap', overflow:'hidden', textOverflow:'ellipsis' }}}
              >
                { this.state.selectedmenu!='general' ?
                  powertoys_dict[this.state.selectedmenu].name + " Settings" :
                  "General Settings"
                }
              </Text>
            </div>
            <div className='editorheadbuttons'>
              <Stack horizontal={true} tokens={{childrenGap:16}}>
                <PrimaryButton
                  text='Save'
                  onClick={saveClicked}
                  />
                <DefaultButton
                  text='Discard'
                  onClick={discardChanges}
                  />
                <IconButton iconProps={{iconName:'Cancel'}} title='Close' />
              </Stack>
            </div>
          </div>
          <div className='editorbody'>
            <ScrollablePane
            styles= {{
              contentContainer: {
                padding: '30px',
              }
            }}
            >
            {
              (() => {
                if(this.state.selectedmenu === 'general') {
                  return <GeneralSettings
                    key="general"
                    settings={this.state.settings.general}
                    ref={(input:any) => {this.settingsscreenref = input;}}
                  />
                } else if(this.state.selectedmenu in this.state.settings.powertoys) {
                  return <CustomSettingsScreen
                    key={this.state.selectedmenu}
                    powertoy={this.state.settings.powertoys[this.state.selectedmenu]}
                    ref={(input:any) => {this.settingsscreenref = input;}}
                    />
                } else {
                  return null;
                }
              })()
            }
            </ScrollablePane>
          </div>
        </div>
      </div>
    );
  }
};
