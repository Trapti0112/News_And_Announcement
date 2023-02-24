import * as React from 'react';
import styles from './Header.module.scss';
import { IHeaderProps } from './IHeaderProps';

export default class Header extends React.Component<IHeaderProps, {}> {
    public render(): React.ReactElement<IHeaderProps> {
        return(
            <div className={styles.header}>{/**className="col-sm-12 col-lg-12" */}
            <div className={styles.sp_dashboard_compnent}>
              <div className={styles.heddings_copm}>
                <h4>
                  {this.props.Title}
                </h4>
                {this.props.showSeAllLink && (<a href={this.props.seeAllLink} target="_blank" className={styles.seeAllLink} data-interception="off">See all</a>)}
                {/* <a href="#">See All</a> */}
              </div>
            </div>
          </div>
        )
    }
}