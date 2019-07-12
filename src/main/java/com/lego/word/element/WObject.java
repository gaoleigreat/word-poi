package com.lego.word.element;

public abstract class WObject {

    public abstract Object getValByKey(String key);

    public boolean isPic(String key) {
        return this.getValByKey(key) instanceof WPic;
    }

    public String getTextByKey(String key) {
        if (!isPic(key)) {
            return (String) this.getValByKey(key);
        } else {
            return null;
        }
    }

    public WPic getPicByKey(String key) {
        if (isPic(key)) {
            return (WPic) this.getValByKey(key);
        } else {
            return null;
        }
    }
}
