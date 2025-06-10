package com.ind;

public record StyleModel(String id, String name, String font, String type, String color, int fontSize,String alignment, double paragraphSpacing, double lineSpacing, double paragraphBeforeSpacing, boolean bold) {
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

    public String getAlignment() {
        return alignment;
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

    public boolean isBold() {
        return bold;
    }
}