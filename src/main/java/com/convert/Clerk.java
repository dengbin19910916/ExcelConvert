package com.convert;

import com.convert.annotation.Sheet;
import com.convert.annotation.Title;
import com.convert.annotation.Workbook;

import java.time.LocalDate;

/**
 * 行员
 *
 * Created by Dengbin on 2015/10/3.
 */
@Workbook(
        name = "行员管理",
        sheets = {
                @Sheet(name = "城西支行", mapper = Clerk.class),
                @Sheet(name = "城东支行", mapper = Clerk.class)
        },
        type = Workbook.Type.XLSX
)
public class Clerk {

    @Title(name = "行员代号")
    private String id;
    @Title(name = "行员名称")
    private String name;
    @Title(name = "出生日期")
//    private String birthday;
    private LocalDate birthday;
    @Title(name = "业务量")
//    private String portfolio;
    private int portfolio;
    @Title(name = "绩效")
//    private String performance;
    private int performance;
    @Title(name = "发放日期")
    private String issueDate;
    @Title(name = "备注")
    private String comment;

    public Clerk() {
        super();
    }

    public Clerk(String id, String name, LocalDate birthday) {
        this.id = id;
        this.name = name;
        this.birthday = birthday;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public LocalDate getBirthday() {
        return birthday;
    }

    public void setBirthday(LocalDate birthday) {
        this.birthday = birthday;
    }

    public int getPortfolio() {
        return portfolio;
    }

    public void setPortfolio(int portfolio) {
        this.portfolio = portfolio;
    }

    public int getPerformance() {
        return performance;
    }

    public void setPerformance(int performance) {
        this.performance = performance;
    }

    public String getIssueDate() {
        return issueDate;
    }

    public void setIssueDate(String issueDate) {
        this.issueDate = issueDate;
    }

    public String getComment() {
        return comment;
    }

    public void setComment(String comment) {
        this.comment = comment;
    }

    @Override
    public String toString() {
        return "Clerk{" +
                "id='" + id + '\'' +
                ", name='" + name + '\'' +
                ", birthday='" + birthday + '\'' +
                ", portfolio='" + portfolio + '\'' +
                ", performance='" + performance + '\'' +
                ", issueDate='" + issueDate + '\'' +
                ", comment='" + comment + '\'' +
                '}';
    }
}
