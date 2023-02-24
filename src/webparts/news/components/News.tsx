import * as React from 'react';
import styles from './News.module.scss';
import { INewsProps } from './INewsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import Header from './header/Header';
import spservices from '../../../spservices/spservices';
// Import css files //npm i react-slick and npm install slick-carousel
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import Slider from "react-slick";//npm i --save-dev @types/react-slick
import { INewsStats } from './INewsStats';

export default class News extends React.Component<INewsProps, INewsStats> {

  private spService: spservices;
  private nItems: any[] = [];
  private allNews: any[] = [];

  constructor(props: any) {
    super(props);
    this.state = {
      newsItems: [],
      devisionCounts: [],
      isLoading: true,
      choices: [],
      filterOption: "All",
      isSelectLoading: true,
      isListNameCorrect: false,
      moreNewsPageLink:""
    };
    this.spService = new spservices(this.context);
  }
  componentDidMount = () => {
    let seeAllPageURL=this.props.moreNewsPageUrl + "?listName=" + this.props.listName;
    this.setState({
      moreNewsPageLink:seeAllPageURL
    });
    this._renderLatestCompanyNewsDataAsync("All");
  };

  private _renderLatestCompanyNewsDataAsync(onSelect: any) {
    this.setState({
      isLoading: true,
      filterOption: onSelect
    });
    this._getLatestNewsData(onSelect).then((Response:any) => {
      // console.log(Response);
      
      this.setState({ isLoading: false });
      this._renderListCompanyNews(Response);
    });
  }

  private _renderListCompanyNews(itemsNews: any[]): void {
    this.nItems = itemsNews;
    this.setState({ newsItems: itemsNews });
  }


  private async _getLatestNewsData(onSelect: any): Promise<any[]> {
    var webPartData:any = [];
    let toDay: any = new Date();
    let date = new Date(toDay).toLocaleDateString(this.props.context.pageContext.cultureInfo.currentCultureName);//Division,IsRedirectUrl //ThumbnailImageUrl, 
    let selectQuery = "ID, Title, ThumbnailImage, ShortDescription, Description, PublishDate,RedirectUrl, IsActive, ExpiryDate, OpenLinkInNewTab";
    let filterQuery = "";
    if (onSelect == "All") {
      filterQuery =  "(IsActive eq 1 and PublishDate le '" + date + "' and (ExpiryDate ge '" + date + "' or ExpiryDate eq null) )";
    }
    // else {
    //   filterQuery = `IsActive eq 1 and PublishDate le '${new Date().toISOString()}'`;
    // }

    const data = await this.spService.getListItems(this.props.context.pageContext.site.absoluteUrl, this.props.listName, selectQuery, "", filterQuery, 12, this.props.orderBy, this.props.order == "Ascending" ? true : false).then((item:any) => {
      // console.log(item)
      this.allNews = item;
      webPartData = item.slice(0, this.props.itemsLimit);
      this.setState({ newsItems: item, isListNameCorrect: true  });
      return webPartData;
    },((err:any) => {
      if(err.response.status == 404) {
        this.setState({isListNameCorrect: false, isLoading: false});
      }
      console.log(err);
    }));
    return webPartData;
  }

  public render(): React.ReactElement<INewsProps> {
   const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    let slideToShow = 3;
    let slideToScroll = 3;
    if(window.innerWidth <= 760) {
      slideToShow = 1;
      slideToScroll = 1;
    }
    else if(this.state.newsItems!=undefined && this.state.newsItems!=null && this.state.newsItems.length!=0 && this.state.newsItems.length < 3) {
      slideToShow = this.state.newsItems.length;
      slideToScroll = this.state.newsItems.length;
    }
    let settings = {
      dots: true,
      infinite: true,
      speed: 500,
      slidesToShow: slideToShow,
      slidesToScroll: slideToScroll,
      //slidesToShow: (this.state.newsItems!=undefined && this.state.newsItems!=null && this.state.newsItems.length!=0 && this.state.newsItems.length < 3)? this.state.newsItems.length: 3,
      //slidesToScroll: (this.state.newsItems!=undefined && this.state.newsItems!=null && this.state.newsItems.length!=0 && this.state.newsItems.length < 3)? this.state.newsItems.length: 3,
      autoplay: true,
      autoplaySpeed: (this.props.speedOfCarousel * 1000),
      arrows: false
    };
    let heightOfComponent = parseInt(this.props.SetHeight) - 60;
    let showSeAllLink = (this.state.isListNameCorrect == false && this.state.newsItems.length == 0)? false : true;
    
    return (
      <section className={`${styles.news}`}>
       <Header Title={this.props.webpartTitle} seeAllLink={this.state.moreNewsPageLink} showSeAllLink={showSeAllLink}></Header>
        {this.state.isLoading ?
        <ProgressIndicator label="Please wait..." /> :
        <div style={{ height: `${heightOfComponent}px`}} className={styles.newsBody}>
          {this.state.isListNameCorrect == false && this.state.newsItems.length == 0 && (
            <p className={styles.errorMessage}>Please update the list name in property pane. </p>
          )}
          {(this.state.newsItems.length == 0 && this.state.isListNameCorrect == true) && (
            <p className={styles.errorMessage}>{this.props.emptyData}</p>) 
          }
          {(this.state.newsItems.length !== 0 && this.state.isListNameCorrect == true) && (
            <Slider {...settings}>
              {/* {arNewsItems} */}
              {this.state.newsItems.length > 0 && this.state.newsItems.map((item:any) =>{
                let imageUrl = "";
                let redirectUrl = "";
                let target:string = "_self";
                // if (item.ThumbnailImageUrl!=undefined && item.ThumbnailImageUrl!=null && item.ThumbnailImageUrl["Url"]!=undefined && item.ThumbnailImageUrl["Url"]!=null) {
                //   imageUrl = item.ThumbnailImageUrl["Url"];
                // }
                let newsThumbnailData = JSON.parse(item["ThumbnailImage"]);
                if(newsThumbnailData!=undefined && newsThumbnailData!=null){
                  imageUrl = newsThumbnailData["serverRelativeUrl"];
                }
                else if(this.props.defaultThumbnail!=undefined && this.props.defaultThumbnail!=null && this.props.defaultThumbnail!="") {
                  imageUrl = this.props.defaultThumbnail;
                }
                else {
                  imageUrl = String(require('../assets/Logo.png'));
                }
                //let newsThumbnailData = JSON.parse(responseJSON.value[i]["Image"]);
                  
                if (item.RedirectUrl!=undefined && item.RedirectUrl!=null && item.RedirectUrl["Url"]!=undefined && item.RedirectUrl["Url"]!=null) {
                  redirectUrl = item.RedirectUrl["Url"];
                }
                if (redirectUrl == "") {//&& (item["IsRedirectUrl"] == true)
                  redirectUrl = this.props.newsDetailsPageUrl + "?newsid=" + item.ID + "&listName=" + this.props.listName;
                }
                if(item.OpenLinkInNewTab){
                  target = "_blank";
                }
                
                return (
                  <div className={styles.newsItem}>
                    <div className={styles.imageDiv} style={{ height: `${heightOfComponent/2}px`}}><img src={imageUrl}></img></div>
                    <h4><a href={redirectUrl} target={target} data-interception="off" title={item.Title}>{item.Title}</a></h4>
                    <p title={item.ShortDescription}>{item.ShortDescription}</p>
                  </div>
                )
              })}
              </Slider>
          )}
          </div>
        }
      </section>
    );
  }
}
