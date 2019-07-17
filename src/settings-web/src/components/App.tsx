import React from 'react';
import {Stack, Text, Nav, PrimaryButton, ScrollablePane, INavLink, Spinner, SpinnerSize} from 'office-ui-fabric-react';
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
      settings: {}
    }
  }

  public componentDidMount() {
    this.send_message_to_application(JSON.stringify({'refresh':true}));
  }

  public send_message_to_application(msg: string) {
    (window as any).output_from_webview(msg);
  }

  public receive_config_msg(config: any):void {
    let current_selected_menu = this.state.selectedmenu;
    if(!config.hasOwnProperty('powertoys') || !config.powertoys.hasOwnProperty(current_selected_menu)) {
      current_selected_menu='general';
    }
    this.setState({settings: config, selectedmenu: current_selected_menu});
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
      // output_from_webview should be declared in index.html
      (window as any).output_from_webview(JSON.stringify(this.settingsscreenref.get_data()));
    };


    return (
      <div className='body'>
        <div className='sidebar'>
          <Nav
            selectedKey= {this.state.selectedmenu}
            onLinkClick = {
              (ev?: React.MouseEvent<HTMLElement,MouseEvent>, item?: INavLink) => {
                this.setState({selectedmenu : ((item && item.key)||null) });
              }
            }
            styles = {{
              navItems: { margin : '0'},
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
                links: [
                  { name: 'General Settings', key:'general', url:'', icon: 'Settings' }
                ].
                concat(powertoys_links)
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
                  "PowerToys General Settings"
                }
              </Text>
            </div>
            <div className='editorheadbuttons'>
              <Stack horizontal={true} tokens={{childrenGap:16}}>
                <PrimaryButton
                  text='Save'
                  onClick={saveClicked}
                  />
              </Stack>
            </div>
          </div>
          <div className='editorbody'>
            <ScrollablePane
            styles= {{
              contentContainer: {
                paddingTop: '30px',
                paddingLeft: '30px',
                paddingRight: '30px'
                // padding bottom will be applied by an empty span in the contents, for edge compatibility.
              }
            }}
            >
            {
              (() => {
                if(this.state.selectedmenu === 'general' && this.state.settings.hasOwnProperty('general')) {
                  return <GeneralSettings
                    key="general"
                    settings_key="general"
                    settings={this.state.settings.general}
                    ref={(input:any) => {this.settingsscreenref = input;}}
                  />
                } else if( this.state.settings.hasOwnProperty('powertoys') && this.state.selectedmenu in this.state.settings.powertoys) {
                  return <CustomSettingsScreen
                    key={this.state.selectedmenu}
                    settings_key={this.state.selectedmenu}
                    powertoy={this.state.settings.powertoys[this.state.selectedmenu]}
                    ref={(input:any) => {this.settingsscreenref = input;}}
                    /> 
                } else { 
                  return <Spinner size={SpinnerSize.large} label="Loading the Settings..." labelPosition="top" />
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
