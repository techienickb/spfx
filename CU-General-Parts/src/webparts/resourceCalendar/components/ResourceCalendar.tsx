import * as React from 'react';
import styles from './ResourceCalendar.module.scss';
import { Calendar, Views } from 'react-big-calendar';
import localizer from 'react-big-calendar/lib/localizers/moment';
import { Spinner, SpinnerSize, Label, Stack, DatePicker } from 'office-ui-fabric-react';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { MSGraphClient } from '@microsoft/sp-http';
import * as moment from 'moment';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DayOfWeek } from 'office-ui-fabric-react/lib/Calendar';
import * as strings from 'ResourceCalendarWebPartStrings';

interface IEvent {
  id: string|number;
  resourceId: string;
  start: Date;
  end: Date;
  title: string;
}

export interface IResourceCalendarProps {
  resources: IPropertyFieldGroupOrPerson[];
  context: WebPartContext;
  mode: string;
}


export default class ResourceCalendar extends React.Component<IResourceCalendarProps, { events: IEvent[], resources: IPropertyFieldGroupOrPerson[], date: Date  }> {
  private events: IEvent[];
  private loc = localizer(moment);

  constructor(props: IResourceCalendarProps) {
    super(props);
    this.state = { events: null, resources: null, date: new Date() };
  }

  public componentDidMount(): void {
    this.onSelectDate(moment(), moment(), this.props.resources);
  }

  private onSelectDate = (start: Date|moment.Moment, end?: Date|moment.Moment, _res?: IPropertyFieldGroupOrPerson[]):void => {
    const { resources } = _res ? { resources: _res } : this.props;
    this.events = [];
    resources.forEach(r => {
      this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
        var now = moment(start);
        now.hours(0);
        now.minutes(0);
        now.seconds(0);
        if (!end) end = start;
        var _end = moment(end);
        _end.hours(23);
        _end.minutes(59);
        _end.seconds(59);
        client.api(`users/${r.email}/events`).filter(`start/dateTime ge '${now.toJSON()}' and end/dateTime le '${_end.toJSON()}'`).select("subject,start,end,id").get((err, res: any) => {
          if (res == null) {
            client.api(`me/calendar/getschedule`).post({ Schedules: [r.email], StartTime: { dateTime: now.toJSON(), timeZone: now.format('z') }, EndTime: { dateTime: _end.toJSON(), timeZone: _end.format('z') }, availabilityViewInterval: 15 },(err2, res2: any) => {
              if (res2 == null) return;
              let si: MicrosoftGraph.ScheduleInformation[] = res2.value;
              si.forEach(_s => _s.scheduleItems.forEach(_i => this.events.push({ id: _i.subject, resourceId: r.email, start: moment.utc(_i.start.dateTime).toDate(), end: moment.utc(_i.end.dateTime).toDate(), title: `${_i.subject ? _i.subject : ''} ${_i.status} ${_i.location ? _i.location : '' }` })));
              this.setState({ ...this.state, events: this.events });
            });
          } else {
            let ev: MicrosoftGraph.Event[] = res.value;
            ev.forEach(e => this.events.push({ id: r.id, resourceId: r.email, start: moment.utc(e.start.dateTime).toDate(), end: moment.utc(e.end.dateTime).toDate(), title: e.subject }));
            this.setState({ ...this.state, events: this.events });
          }
        });
      });
    });
  }

  private onRangeChange = (dates: Date[]|moment.Moment[]): void => {
    this.onSelectDate(dates[0], dates[dates.length - 1]);
  }

  public componentWillReceiveProps(nextProps: Readonly<IResourceCalendarProps>, nextContext: any): void {
    this.onSelectDate(moment(), moment(), nextProps.resources);
  }

  private selectDate = (date: Date, selectedDateRangeArray?: Date[]): void => {
    this.setState({...this.state, date: date });
    this.onSelectDate(date);
  }

  public render(): React.ReactElement<IResourceCalendarProps> {
    const { events,date } = this.state;
    const { resources, mode }  = this.props;
    const _24: number[] = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23];

    return (
      <div className={ styles.resourceCalendar }>
        <div className={ styles.container }>
          {(resources == null || resources.length == 0) && <Label>Select some resources in the properties pane</Label>}
          {events == null && resources.length > 0 && <Spinner size={SpinnerSize.large} />}
          {events != null && mode !== "horizontal" && <Calendar localizer={this.loc} onRangeChange={this.onRangeChange} events={events} defaultView={Views.DAY} culture="en-GB" views={['day', 'work_week']} 
            step={60} defaultDate={new Date()} resources={resources.map(r => ({ resourceId: r.email, resourceTitle: r.fullName }))} resourceIdAccessor="resourceId" resourceTitleAccessor="resourceTitle" />}
          {events != null && mode === "horizontal" && <div>
            <div style={{padding: '4px 0'}}>
              <DatePicker onSelectDate={this.selectDate} value={date} strings={strings.calStrings} isMonthPickerVisible={true} firstDayOfWeek={DayOfWeek.Monday}  />
            </div>
            <div className={styles.horz}>
              <header>
                <div>Resource</div>
                {resources.map(r => (<div>{r.fullName}</div>))}
              </header>
              <article>
                <div>
                  {_24.map(i => (<span style={ i === 0 ? null : { gridColumn: `${(i * 4) + 1} / span 4`}}>{i}:00</span>))}
                </div>
                {resources.map(r => (<div>
                  { events.filter(e => e.resourceId === r.email).map(e => {
                    const g = new Date(e.end.getTime() - e.start.getTime());
                    let c = (g.getHours() - 1) * 4 + Math.floor(g.getMinutes() / 15);
                    return (<span title={e.title} style={{ gridColumn: `${(e.start.getHours() * 4) + Math.floor(e.start.getMinutes() / 15) + 1} / span ${c}` }}>{e.title}</span>);
                  }) }
                  </div>))}
              </article>
            </div>
          </div>}
        </div>
      </div>
    );
  }
}