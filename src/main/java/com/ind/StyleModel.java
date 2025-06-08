package com.ind;

public record StyleModel(String id, String name, String font, String type, String color, int fontSize, double paragraphSpacing, double lineSpacing, double paragraphBeforeSpacing) {
    public String getId(){
        return id;
    }
    public String getName(){ 
        return name; 
    }
    
    public String getFont() {
        return font;
    }

    public String getType() {
        return type;
    }

    public String getColor() {
        return color;
    }

    public int getFontSize() {
        return fontSize;
    }

    public double getParagraphSpacing() {
        return paragraphSpacing;
    }

    public double getLineSpacing() {
        return lineSpacing;
    }
    
    public double getParagraphBeforeSpacing() {
        return paragraphBeforeSpacing;
    }
}