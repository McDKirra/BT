import * as React from 'react';
import styles from './Bot.module.scss';
import { IBotProps } from './IBotProps';
import ReactWebChat from 'botframework-webchat';
import { DirectLine } from 'botframework-directlinejs';
import { sp } from '@pnp/sp';

export interface IBotState {
  directLink: any;
}

export default class Bot extends React.Component<IBotProps, IBotState> {
  constructor(props) {
    super(props);

    this.state = {
      directLink: new DirectLine({
        secret: "Ihx07DHSM8w.-ZSO_6YTtYu3qXVWoKWCPcWGiGKl9NQtGwz5peB1QdQ"
      })
    };
  }

  public render(): React.ReactElement<IBotProps> {

    //Current User information from Context
    var user = { id: this.props.context.pageContext.user.email, name: this.props.context.pageContext.user.displayName };

    //Sending BOT "event" type dialog with user basic information for greeting.
    this.state.directLink
      .postActivity({ type: "event", name: "sendUserInfo", value: this.props.context.pageContext.user.displayName, from: user })
      .subscribe(id => console.log("success", id));

    //Subscribing for activities created by BOT
    var act: any = this.state.directLink.activity$;
    act.subscribe(
      a => {debugger;
        if (a.type == "event" && a.name == "search") {
          sp.search({
            Querytext: "filetype:pdf"
          }).then((results) => {
            var items = results.PrimarySearchResults.map((value) => {
              return {
                "Title": value.Title,
                "FileExtension": value.FileExtension,
                "Author": value.Author,
                "SubText": value.HitHighlightedSummary,
                "PictureUrl": value.PictureThumbnailURL,
                "redirectUrl": value.ParentLink
              }
            })
            this.state.directLink
              .postActivity({ type: "message", text: "showresults", value: items, from: user })
              .subscribe(id => { console.log("success", id) });
          })
        }
      }
    );

    return (
      <div className={styles.bot} style={{ height: 700 }}>
        <ReactWebChat directLine={this.state.directLink}  />
      </div>
    );
  }
}
