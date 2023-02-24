import * as React from 'react';
import styles from './AllNews.module.scss';
import { IAllNewsProps } from './IAllNewsProps';
import { IAllNewsStats } from './IAllNewsStats';
import { escape } from '@microsoft/sp-lodash-subset';
import spservices from '../../../spservices/spservices';
import spcommon from '../../../spservices/spcommon';
import * as moment from 'moment';
import * as dateformat from 'dateformat';  //npm i dateformat and npm i --save-dev @types/dateformat
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { Button, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';

let listName ="";

export default class AllNews extends React.Component<IAllNewsProps, IAllNewsStats> {
  private spService: spservices = null;
  private allNews: any[] = [];
  private yearMonthCount: any = [];

  public constructor(props: any) {
    super(props);
    this.state = {
      newsItems: [],
      divisions: [],
      selectedDivision: "All",
      selectedYear: "",
      selectedMonth: "",
      searchText: "",
      isLoading: true,
      listName:""
    };
    this.spService = new spservices(this.context);
  }

  resetSearch = () =>{
    this._getYearMonthCount(this.allNews);
    this.setState({ newsItems: this.allNews, selectedYear: "", selectedMonth: "", searchText: "" });
  }

  filteredNewsBySearch = (value:any) =>{
    if(value!=""){
      let filteredNews: any[] = this.filterNewsBySearch(value, this.allNews);
      this._getYearMonthCount(filteredNews);
      this.setState({ newsItems: filteredNews, selectedYear: "", selectedMonth: ""});
    }
  }

  searchTextValue = (value:any) =>{
    this.setState({searchText: value});
  }

  private filterNewsBySearch(searchText: any, filteredNews: any[]): any[] {
    return filteredNews.filter(newsItem => {
      //let dateP: any = moment(newsItem.PublishDate).format("MM-DD-YYYY");
      let dateP: any = moment(newsItem.PublishDate).format('ll');
      let itemfound = false;
      if (searchText != "" && (newsItem.Title.toLowerCase().indexOf(searchText.toLowerCase()) >= 0 || dateP.toLowerCase().indexOf(searchText.toLowerCase()) >= 0 || (newsItem.ShortDescription != null && newsItem.ShortDescription.toLowerCase().indexOf(searchText.toLowerCase()) >= 0) || (newsItem.Description != null && newsItem.Description.toLowerCase().indexOf(searchText.toLowerCase()) >= 0))) {
        itemfound = true;
      }
      return itemfound;
    });
  }

  private _getNewsByMonth(month: any, year: any) {
    let filteredNews: any[] = this.allNews;
    if (this.state.searchText != ""){
      filteredNews = this.filterNewsBySearch(this.state.searchText, filteredNews);
    }
    filteredNews = this.filterNewsByMonth(month, year, filteredNews);

    this.setState({
      newsItems: filteredNews,
      selectedYear: year,
      selectedMonth: month
    });
  }

  private filterNewsByMonth(month: any, year: any, filteredNews: any[]): any[] {
    return filteredNews.filter(newsItem => {
      let pYear: any = moment(newsItem.PublishDate).format("YYYY");
      let pMonth: any = moment(newsItem.PublishDate).format("MMMM");
      return pYear == year && pMonth == month;
    });
  }

  private _getNewsByYear(year: any) {
    let filteredNews: any[] = this.allNews;
    if (this.state.searchText != "")
      filteredNews = this.filterNewsBySearch(this.state.searchText, filteredNews);

    if (year != this.state.selectedYear)
      filteredNews = this.filterNewsByYear(year, filteredNews);
    else
      year = "";

    this.setState({
      newsItems: filteredNews,
      selectedYear: year,
      selectedMonth: ""
    });
  }

  private filterNewsByYear(year: any, filteredNews: any[]): any[] {
    return filteredNews.filter(newsItem => {
      let pYear: any = moment(newsItem.PublishDate).format("YYYY");
      return pYear == year;
    });
  }

  //Get count of items based on particular month of particular year
  private _getYearMonthCount(filteredNews: any[]) {
    this.yearMonthCount = [];
    filteredNews.forEach(newsItem => {
      let pYear: any = moment(newsItem.PublishDate).format("YYYY");
      let pMonth: any = moment(newsItem.PublishDate).format("MMMM");
      if (this.yearMonthCount.length > 0) {
        let yrF = false;

        for (let yr of this.yearMonthCount) {
          if (yr.title == pYear) {//PublishDate Year entry is already present
            yrF = true;
            let mnthF = false;
            for (let mnth of yr.months) {
              if (mnth.title == pMonth) {//PublishDate Month entry is already present
                mnthF = true;
                mnth.count++;//Increase count for month
              }
            }
            if (!mnthF) {//PublishDate Year entry is there but month entry is not present so just add that entry
              yr.months.push({
                title: pMonth,
                count: 1
              });
            }
          }
        }
        if (!yrF) {//If PublishDate year entry is not present so add that
          this.yearMonthCount.push({
            title: pYear,
            months: [
              {
                title: pMonth,
                count: 1
              }
            ]
          });
        }
      } else {//If PublishDate year and month entry is not present, First item
        this.yearMonthCount.push({
          title: pYear,
          months: [
            {
              title: pMonth,
              count: 1
            }
          ]
        });
      }
    });
  }

  public async componentDidMount() {
    const urlParams = new URLSearchParams(window.location.search);
    listName = urlParams.get('listName');
    // const root = document.documentElement;
    // root?.style.setProperty(
    //   "--font-family","Montserrat"
    // );
    this._renderListCompanyNewsDataAsync();
  }

  private _renderListCompanyNewsDataAsync() {
    this._getListComponyNewsData().then((newsData) => {
      this.allNews = newsData;
      this._getYearMonthCount(this.allNews);
      this.setState({
        newsItems: this.allNews,
        selectedYear: "",
        selectedMonth: "",
        searchText: "",
        isLoading: false
      });
    });
  }

  private async _getListComponyNewsData(): Promise<any[]> {//ThumbnailImageUrl, 
    let selectQuery = "ID, Title, ThumbnailImage, ShortDescription, Description, PublishDate, IsActive,RedirectUrl, ExpiryDate, OpenLinkInNewTab";
    let toDay: any = new Date();
    let date = new Date(toDay).toLocaleDateString(this.props.context.pageContext.cultureInfo.currentCultureName);
    let filterQuery = "(IsActive eq 1 and PublishDate le '" + date + "' and (ExpiryDate ge '" + date + "' or ExpiryDate eq null) )";
    const data = await this.spService.getListItems(this.props.context.pageContext.site.absoluteUrl, listName, selectQuery, "", filterQuery, 4999, "PublishDate", false);
    return data;
  }

  public render(): React.ReactElement<IAllNewsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    let arNewsItems:any = [];
    this.state.newsItems.forEach((item: any, itemKey) => {
      let title = "";
      let imageUrl = "";
      let newsShortDescription = "";
      let publishDate = "";
      let redirectUrl = "";

      if (!spcommon.fIsNullOrUndefined(item.Title, false)) {
        title = item.Title;
      }

      // if (!spcommon.fIsNullOrUndefined(item.ThumbnailImageUrl, false) && !spcommon.fIsNullOrUndefined(item.ThumbnailImageUrl["Url"], false)) {
      //   imageUrl = item.ThumbnailImageUrl["Url"];
      // } else if(!spcommon.fIsNullOrUndefined(this.props.DefaultThumbnail, false) && this.props.DefaultThumbnail!="") {
      //   imageUrl = this.props.DefaultThumbnail;
      // }
      let newsThumbnailData = JSON.parse(item["ThumbnailImage"]);
      if(!spcommon.fIsNullOrUndefined(newsThumbnailData, false)){
        imageUrl = newsThumbnailData["serverRelativeUrl"];
      }
      else if(!spcommon.fIsNullOrUndefined(this.props.DefaultThumbnail, false) && this.props.DefaultThumbnail!="") {
        imageUrl = this.props.DefaultThumbnail;
      }
      else
      //imageUrl=this.props.DefaultThumbnail;
      //imageUrl = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTg7evYs1TlJ6nC62yZVs7bp_r-4u_S__aYmQylI-UPFg&s";
        imageUrl = String(require('../assets/Logo.png'));
        //('../../../assets/images/noimagerectangle.jpg')
        // imageUrl = `${this.props.siteUrl}/Company News Images/Divisions Logo/${item.Division}.png`;
      
      if (!spcommon.fIsNullOrUndefined(item.PublishDate, false) && item.PublishDate !== "") {
        publishDate = moment(item.PublishDate).format('ll');
      }
      if (!spcommon.fIsNullOrUndefined(item.ShortDescription, false) && item.ShortDescription !== "") {
        newsShortDescription = item.ShortDescription;
      }

      if (!spcommon.fIsNullOrUndefined(item.RedirectUrl, false) && !spcommon.fIsNullOrUndefined(item.RedirectUrl["Url"], false)) {
        redirectUrl = item.RedirectUrl["Url"];
      }

      let target: string = "_self";
      if (item.OpenLinkInNewTab)
        target = "_blank";

      if (redirectUrl == "") {
     
        redirectUrl = this.props.newsDetailsPageUrl + "?newsid=" + item.ID + "&listName=" + listName;

      }
      arNewsItems.push(
        <div className={styles.moreNewsItem}>
          <div className={styles.moreNewsImg}>
            <a href={redirectUrl} target={target} data-interception="off" title={title}> <img src={imageUrl} /></a> </div>
          <div className={styles.moreNewsContent}>
            <h4><a href={redirectUrl} target={target} title={title} data-interception="off"> {title}</a></h4>
            <div className={styles.clearfix}></div>
            <div className={styles.moreDate}>
              {/* { moment(publishDate).format('YYYY-MM-DD')} */}
              {publishDate}
            </div>
            <p>{newsShortDescription}</p>
          </div>
        </div>
      );
    });

    let itemsL = [];
    for (let yM of this.yearMonthCount) {
      let getNewsByYear = () => {
        this._getNewsByYear(yM.title);
      };

      if (this.state.selectedYear == yM.title) {
        itemsL.push(
          <li className={styles.filterYearSelected} onClick={getNewsByYear}>
            <b> <Icon iconName="ChevronDown" /> </b>
            <a href="javascript:void(0)"> <b>{yM.title}</b></a>
          </li>
        );
      } else {
        itemsL.push(
          <li className={styles.filterYear} onClick={getNewsByYear}>
            <b> <Icon iconName="ChevronRight" /> </b>
            <a href="javascript:void(0)"> <b>{yM.title}</b></a>
          </li>
        );
      }
      if (this.state.selectedYear == yM.title) {
        for (let month of yM.months) {
          let getNewsByMonth = () => {
            this._getNewsByMonth(month.title, yM.title);
          };
          if (this.state.selectedMonth == month.title) {
            itemsL.push(
              <li>
                <a onClick={getNewsByMonth}><b>{month.title} ({month.count})</b></a>
              </li>
            );
          } else {
            itemsL.push(
              <li>
                <a onClick={getNewsByMonth}>{month.title} ({month.count})</a>
              </li>
            );
          }
        }
      }
    }

    return (
      <section className={`${styles.allNews} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.mainContainer}>
          {/* <div className={styles.sp_dashboard_compnent}>
            <div className={styles.heddings_copm}>
              <h4>{this.props.description}</h4>
            </div>
          </div> */}
          <div className={styles.moreNewsContainer}>
            <div className={styles.sidebar}>
              <h4>Year and Month Filter</h4>
              <ul className={styles.filters}>
                {itemsL}
              </ul>
            </div>
            <div className={styles.companyNewsRight}>
              <div className={`ms-Grid ${styles.searchBox}`}>
                <div className={`ms-Grid-row ${styles.msGridRow} `}>
                  <div className={`ms-Grid-col ms-u-sm8 ${styles.msGridCol}`}>
                    <SearchBox
                      className={`react-search-box`}
                      onSearch={newValue => this.searchTextValue(newValue)}
                      onChange={(event?: React.ChangeEvent<HTMLInputElement>, newValue?: string) => this.searchTextValue(newValue)}
                      value={this.state.searchText} placeholder="Search news">
                    </SearchBox>
                  </div>
                  <div className="ms-Grid-col">{/** ms-u-sm1 */}
                    {/* <Button id="SearchButton" onClick={()=>this.filteredNewsBySearch(this.state.searchText)} className={styles.srchBtn}>
                      Search
                    </Button> */}
                    <PrimaryButton onClick={()=>this.filteredNewsBySearch(this.state.searchText)} className={styles.srchBtn}>Search</PrimaryButton>
                  </div>
                  <div className="ms-Grid-col">{/** ms-u-sm1 */}
                    {/* <Button id="ResetButton" onClick={this.resetSearch} className={styles.rsetBtn}>
                      Reset
                    </Button> */}
                    <PrimaryButton onClick={this.resetSearch} className={styles.rsetBtn}>Reset</PrimaryButton>
                  </div>
                </div>
              </div>
              {this.state.isLoading ?
                <ProgressIndicator label="Please wait..." /> :
                // <Scrollbars style={{ width: "100%", height: "235px" }} autoHide={true} renderTrackHorizontal={props => <div {...props} style={{ display: 'none' }} className="track-horizontal" />} >
                  <div className={styles.moreNewsPage}>
                    {arNewsItems}
                    {arNewsItems.length == 0 ? <div className={styles.moreNewsPageNoRecordFound}>No record(s) found.</div> : ''}
                  </div>
                // </Scrollbars>
              }
            </div>
            <div className={styles.clearfix}></div>
          </div>
        </div>
      </section>
    );
  }
}
