import { Injectable, Inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable, BehaviorSubject } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class ReportService {

  constructor(private http: HttpClient, @Inject('BASE_URL') private baseUrl: string) {}

  processReport(templateName: string): Observable<any>
  {
    return this.http.get(this.baseUrl + 'reportrunner?templateName=' + templateName);
  }

  getTemplateTypes(): Observable<any> {
    return this.http.get(this.baseUrl + 'reportrunner/types');
  }

}
