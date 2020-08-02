package com.github.qrpcode.domain;

import java.util.Date;

/**
 * 生成文档信息
 * @author QiRuipeng
 */
public class WordCore {

    private String title;
    private String subject;
    private String creator;
    private String keywords;
    private String description;
    private String lastModifiedBy;
    private int revision;
    private Date created;
    private Date modified;

    public WordCore() {
    }

    public WordCore(String title) {
        this.title = title;
    }

    public WordCore(String title, String creator) {
        this.title = title;
        this.creator = creator;
    }

    public WordCore(String title, String subject, String creator,
                    String keywords, String description, String lastModifiedBy,
                    int revision, Date created, Date modified) {
        this.title = title;
        this.subject = subject;
        this.creator = creator;
        this.keywords = keywords;
        this.description = description;
        this.lastModifiedBy = lastModifiedBy;
        this.revision = revision;
        this.created = created;
        this.modified = modified;
    }

    public String getTitle() {
        return this.title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getSubject() {
        return subject;
    }

    public void setSubject(String subject) {
        this.subject = subject;
    }

    public String getCreator() {
        return creator;
    }

    public void setCreator(String creator) {
        this.creator = creator;
    }

    public String getKeywords() {
        return keywords;
    }

    public void setKeywords(String keywords) {
        this.keywords = keywords;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public String getLastModifiedBy() {
        return lastModifiedBy;
    }

    public void setLastModifiedBy(String lastModifiedBy) {
        this.lastModifiedBy = lastModifiedBy;
    }

    public int getRevision() {
        return revision;
    }

    public void setRevision(int revision) {
        this.revision = revision;
    }

    public Date getCreated() {
        return created;
    }

    public void setCreated(Date created) {
        this.created = created;
    }

    public Date getModified() {
        return modified;
    }

    public void setModified(Date modified) {
        this.modified = modified;
    }
}
