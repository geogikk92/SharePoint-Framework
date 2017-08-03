import * as React from 'react';
import styles from './AutoComplete.module.scss';
import { IAutoCompleteProps } from './IAutoCompleteProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class AutoComplete extends React.Component<IAutoCompleteProps, void> {
  public render(): React.ReactElement<IAutoCompleteProps> {
    return (
      <div className={styles.autoComplete}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p className="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p className="ms-font-l ms-fontColor-white">{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more!</span>
              </a>
              <div id="lists">
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
