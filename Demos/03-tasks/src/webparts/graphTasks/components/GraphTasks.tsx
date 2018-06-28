import * as React from 'react';
import styles from './GraphTasks.module.scss';
import { IGraphTasksProps } from './IGraphTasksProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IGraphTasksState } from './IGraphTasksState';

import { MSGraphClient } from '@microsoft/sp-client-preview';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { List } from 'office-ui-fabric-react/lib/List';


export default class GraphTasks extends React.Component<IGraphTasksProps, IGraphTasksState> {

  constructor(props: IGraphTasksProps) {
    super(props);

    this.state = {
      tasks: []
    };
  }

  public componentDidMount(): void {
    this.props.graphClient
      .api('/me/planner/tasks')
      .get((error: any, tasksResponse: any, rawResponse?: any) => {
        console.log('tasksResponse', tasksResponse);
        const plannerTasks: MicrosoftGraph.PlannerTask[] = tasksResponse.value;
        this.setState({ tasks: plannerTasks });
      });
  }

  private _onRenderEventCell(item: MicrosoftGraph.PlannerTask, index: number | undefined): JSX.Element {
    return (
      <div>
        <h3>{item.title}</h3>
        <strong>Due:</strong> {item.dueDateTime}
      </div>
    );
  }

  public render(): React.ReactElement<IGraphTasksProps> {
    return (
      <List items={this.state.tasks} 
            onRenderCell={this._onRenderEventCell} />
    );
  }
}
