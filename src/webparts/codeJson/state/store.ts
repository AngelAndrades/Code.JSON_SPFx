import { Observable, BehaviorSubject } from 'rxjs';
import { distinctUntilChanged, pluck } from 'rxjs/operators';

import { State, INITIAL_STATE, StateKey } from './state';

export class Store {
    private subject = new BehaviorSubject<State>(INITIAL_STATE);

    public get value() {
        return this.subject.value;
    }

    public select<T>(name: StateKey): Observable<T> {
        return this.subject.pipe(
            pluck<State, T>(name),
            distinctUntilChanged<T>()
        );
    }

    public set<T>(name: StateKey, state: T) {
        this.subject.next({
            ...this.value, [name]: state
        });
    }
}