import * as React from 'react';
import styles from './Chatbot.module.scss';
import { IChatbotProps } from './IChatbotProps';
import { FontSizes } from '@uifabric/styling';
import { IChatBotState } from './IChatBotState';
import { escape } from '@microsoft/sp-lodash-subset';
import ReactWebChat, { createDirectLine, createStore } from 'botframework-webchat';
import { getId } from 'office-ui-fabric-react/lib/Utilities';
import styleSetOptions from 'botframework-webchat';
import { DirectLine, } from 'botframework-directlinejs';
import {
  getTheme,
  mergeStyleSets,
  FontWeights,
  ContextualMenu,
  DefaultButton,
  Modal,
  IDragOptions,
  IconButton,
  ActivityItem
} from 'office-ui-fabric-react';

const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch'
  },
  header: [
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      background: '#ffffff',
      color: theme.palette.black,
      display: 'flex',
      fontSize: FontSizes.xLarge,
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      padding: '12px 12px 14px 24px'
    }
  ],
  body: {
    flex: '4 4 auto',
    padding: '0px 0px 0px 0px',
    overflowY: 'hidden',
    selectors: {
      p: {
        margin: '14px 0'
      },
      'p:first-child': {
        marginTop: 0
      },
      'p:last-child': {
        marginBottom: 0
      }
    }
  }
});

const iconButtonStyles = mergeStyleSets({
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px'
  },
  rootHovered: {
    color: theme.palette.neutralDark
  }
});

export default class Chatbot extends React.Component<IChatbotProps, IChatBotState> {
  constructor(props) {
    super(props);
    const styleOptions = {
      botAvatarImage: 'https://predictdata.blob.core.windows.net/boticon/ChatBotIcon.JPG',
      rootHeight: '400px',
      rootWidth: '500px',
      bubbleBackground: '#ffffff',
      suggestedActionBorderColor: '#a28755',
      suggestedActionBorderRadius: '25px',
      suggestedActionTextColor: '#a28755',
      botAvatarBackgroundColor: '#ffffff'
    };
    this.state = {
      showModal: false,
      isDraggable: true,
      directline: null,
      styleSetOptions: styleOptions
    };
    this._showModal = this._showModal.bind(this);
    this._closeModal = this._closeModal.bind(this);
  }
  async componentDidMount() {
    this.fetchToken();
  }

  private _titleId: string = getId('title');
  private _subtitleId: string = getId('subText');
  private _dragOptions: IDragOptions = {
    moveMenuItemText: 'Move',
    closeMenuItemText: 'Close',
    menu: ContextualMenu
  };
  private _showModal = (): void => {
    this.setState({ showModal: true });
    this.fetchToken();
  };

  private _closeModal = (): void => {
    this.setState({ showModal: false });
  };
  async fetchToken() {
    var myToken ='AIXLQPb9Tno.Kw-y4RX4P4Nbzt5un4oDGHBfPnAdmIQFb1NUHWwuA1M';
    const myHeaders = new Headers()
    myHeaders.append('Authorization', 'Bearer ' + myToken)
    const res = await fetch(
      'https://directline.botframework.com/v3/directline/tokens/generate',
      {
        method: 'POST',
        headers: myHeaders
      }
    )
    const { token } = await res.json();
    console.log(token);
    this.setState({
      directline: createDirectLine({ token })
    });
    this.state.directline.postActivity({
      from: { id: "serId", name: "USER_NAME" },
      name: "requestWelcomeDialog",
      type: "event",
      value: "token"
    }).subscribe(
      id => console.log(`Posted activity, assigned ID ${id}`),
      error => console.log(`Error posting activity ${error}`)
    );

  }

  public render(): React.ReactElement<IChatbotProps> {
    return (
      <div className={ styles.chatbot }>
         <DefaultButton style={{ border: 'none', textDecoration: 'none', background: 'none', outline: 'none' }} onClick={this._showModal}>
        <img src="https://predictdata.blob.core.windows.net/boticon/ChatBotIcon.JPG" alt="Chatbot" height="100px" width="100px" />
      </DefaultButton>
      <Modal
        titleAriaId={this._titleId}
        subtitleAriaId={this._subtitleId}
        isOpen={this.state.showModal}
        onDismiss={this._closeModal}
        isBlocking={true}
        containerClassName={contentStyles.container}
        dragOptions={this.state.isDraggable ? this._dragOptions : undefined}
      >
        <div className={contentStyles.header}>
          <span id={this._titleId}>Virtual Assistant</span>
          <IconButton
            styles={iconButtonStyles}
            iconProps={{ iconName: 'Cancel' }}
            ariaLabel="Close popup modal"
            onClick={this._closeModal as any}
          />
        </div>
        <div id={this._subtitleId} className={contentStyles.body}>
          <ReactWebChat directLine={this.state.directline} styleOptions={this.state.styleSetOptions} />
        </div>
      </Modal>
      </div>
    );
  }
}
