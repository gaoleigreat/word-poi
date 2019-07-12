package com.lego.word;

import com.lego.word.element.WObject;
import com.lego.word.element.WPic;

public class MyTable extends WObject {

    private String name;

    private Integer age;
    private WPic bir;

    private String sex;


    @Override
    public Object getValByKey(String key) {


        if (key.equals("name")) {
            return this.name;
        } else if (key.equals("sex")) {
            return this.sex;
        } else if (key.equals("age")) {
            return this.age;
        } else if (key.equals("bir")) {
            return this.bir;
        } else {
            return null;
        }
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public WPic getBir() {
        return bir;
    }

    public void setBir(WPic bir) {
        this.bir = bir;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }
}
