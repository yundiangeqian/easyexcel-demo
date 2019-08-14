package com.cm.easyexcel.util.bean;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.Serializable;

/**
 * @description:
 * @author: caomian
 * @data: 2019/8/13 19:59
 */
public class ExcelTableStyle implements Serializable {
    boolean headIsBold = false;
    short headFontSize = 18;
    String headFontName = "楷体";
    IndexedColors headColor = IndexedColors.BLUE;
    boolean bodyIsBold = false;
    short bodyFontSize = 15;
    String bodyFontName = "隶书";
    IndexedColors bodyColor = IndexedColors.GREEN;

    public boolean isHeadIsBold() {
        return headIsBold;
    }

    public void setHeadIsBold(boolean headIsBold) {
        this.headIsBold = headIsBold;
    }

    public short getHeadFontSize() {
        return headFontSize;
    }

    public void setHeadFontSize(short headFontSize) {
        this.headFontSize = headFontSize;
    }

    public String getHeadFontName() {
        return headFontName;
    }

    public void setHeadFontName(String headFontName) {
        this.headFontName = headFontName;
    }

    public IndexedColors getHeadColor() {
        return headColor;
    }

    public void setHeadColor(IndexedColors headColor) {
        this.headColor = headColor;
    }

    public boolean isBodyIsBold() {
        return bodyIsBold;
    }

    public void setBodyIsBold(boolean bodyIsBold) {
        this.bodyIsBold = bodyIsBold;
    }

    public short getBodyFontSize() {
        return bodyFontSize;
    }

    public void setBodyFontSize(short bodyFontSize) {
        this.bodyFontSize = bodyFontSize;
    }

    public String getBodyFontName() {
        return bodyFontName;
    }

    public void setBodyFontName(String bodyFontName) {
        this.bodyFontName = bodyFontName;
    }

    public IndexedColors getBodyColor() {
        return bodyColor;
    }

    public void setBodyColor(IndexedColors bodyColor) {
        this.bodyColor = bodyColor;
    }
}
