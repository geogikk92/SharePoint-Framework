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
              <span className="ms-font-xl ms-fontColor-white">{escape(this.props.description)}</span>
              <table>
                <tr>
                  <th>Три имена</th>
                  <th>Дни отпуск</th>
                  <th>Длъжност</th>
                  <th>Отдел</th>
                  <th>Направление</th>
                </tr>
                <tr id="trUserInfo">
                </tr>
              </table>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Използвай последнkия отпуск</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
