import { Component, OnInit } from '@angular/core';
import { ReportService } from './report.service';
import * as XLSX from 'xlsx';
import { Observable } from 'rxjs';

type AOA = any[][];

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
})
export class HomeComponent implements OnInit {

  data: any;
  templateType$: Observable<any[]>;
  selectedType : any;

  constructor(private reportService: ReportService) { }

  ngOnInit(): void
  {
    this.templateType$ = this.reportService.getTemplateTypes();
  }

  onTemplateTypeChange() {
    console.log("Template Type: " + this.selectedType)
    this.reportService.processReport(this.selectedType)
      .subscribe(result => {
        var decodedStringAtoB = atob(result);
        const wb: XLSX.WorkBook = XLSX.read(decodedStringAtoB, { type: 'binary' });

        /* grab first sheet */
        const wsname: string = wb.SheetNames[0];
        const ws: XLSX.WorkSheet = wb.Sheets[wsname];

        /* save data */
        this.data = <AOA>(XLSX.utils.sheet_to_json(ws, { header: 1 }));
        console.log(this.data);
      }, error => console.error(error));
  }

}
