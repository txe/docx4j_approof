package com.evo;

import java.util.List;


public class CandidateReportCtx {

    public class RowData {
        public String date;
        public String title;
        public String note;
        public String note2;

        public RowData(String date, String title, String note, String note2) {
            this.date = date;
            this.title = title;
            this.note = note;
            this.note2 = note2;
        }
    }

    public RowData createRowData(String date, String title, String note){
        return new RowData(date, title, note, "");
    }

    public RowData createRowData(String date, String title, String note, String note2){
        return new RowData(date, title, note, note2);
    }

    // methods for DocxStamper
    public List<RowData> getCareerList()     { return CareerList; }
    public List<RowData> getCareerExpList()  { return CareerExpList; }
    public List<RowData> getEduList()        { return EduList; }
    public List<RowData> getEduExpList()     { return EduExpList; }

    public String CandidateName;
    public String CurrentDate;
    public String MaritalStatus;
    public List<RowData> CareerList;
    public List<RowData> CareerExpList;
    public List<RowData> EduList;
    public List<RowData> EduExpList;
}
