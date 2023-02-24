import * as React from 'react';
import styles from './DetailedNews.module.scss';
import { IDetailedNewsProps } from './IDetailedNewsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import spservices from '../../../spservices/spservices';
import spcommon from '../../../spservices/spcommon';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import * as moment from 'moment';

let listName = "";

export default class DetailedNews extends React.Component<IDetailedNewsProps, {newsItem: any, isLoading: boolean,listName:string}> {
  private spService: spservices;
  public constructor(props: any) {
    super(props);
    this.state = {
      newsItem: [],
      isLoading: true,
      listName:""
    };
    this.spService = new spservices(this.context);
  }

  public async componentDidMount() {
    this._renderListCompanyNewsDetails();
  }

  private _renderListCompanyNewsDetails() {
    this._getComponyNewsDataById().then((Response) => {
      this._renderCompanyNewsDetails(Response);
    });
  }

  private async _getComponyNewsDataById(): Promise<any> {

    const urlParams = new URLSearchParams(window.location.search);
    const newsId = urlParams.get('newsid');
    listName = urlParams.get('listName');
    if (!spcommon.fIsNullOrUndefined(newsId, false) && newsId != "") {//, OpenLinkInNewTab //ThumbnailImageUrl,ShortDescription,, IsActive 
      let selectQuery = "ID, Title, Description, PublishDate";
      const newsItemData: any = await this.spService.getListItemByID(this.props.context.pageContext.site.absoluteUrl, listName, selectQuery, parseInt(newsId));

      console.log(newsItemData);
        return newsItemData;

    } else
      return null;
  }

  private _renderCompanyNewsDetails(Response: any) {
    this.setState({ newsItem: Response, isLoading: false });
  }
  public render(): React.ReactElement<IDetailedNewsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.detailedNews} ${hasTeamsContext ? styles.teams : ''}`}>
        {this.state.isLoading ?
          <ProgressIndicator label="Please wait..." /> :
          <div className={styles.innerNews}>
            <div className={styles.innerNewsTitle}>
              <h4>{this.state.newsItem.Title}</h4>
              <p>
                {moment(this.state.newsItem.PublishDate).format('ll')}
              </p>
            </div>
            <div className={styles.innerNewsContent}>
              {/* {likeHtml} */}
              <div dangerouslySetInnerHTML={{ __html: this.state.newsItem.Description }}></div>
            </div>
            { }
          </div>}
      </section>
    );
  }
}
