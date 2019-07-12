package com.lego.word;

import com.lego.word.element.WObject;
import com.lego.word.element.WPic;

public class MyDoc extends WObject {
    @Override
    public Object getValByKey(String key) {

        if (key.equals("name")) {
            return this.name;
        } else if (key.equals("picture")) {
            return this.picture;
        } else if (key.equals("address")) {
            return this.address;
        } else {
            return null;
        }
    }

    private WPic picture;
    private String name;
    private String address;

    public WPic getPicture() {
        return picture;
    }

    public void setPicture(WPic picture) {
        this.picture = picture;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }
}
