from lxml import etree as ET
import zipfile

# XML mnamespaces for the tags we want
ooXMLns = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}


def get_document_comments(docxFileName):
    docx_zip = zipfile.ZipFile(docxFileName)

    comments_xml = docx_zip.read("word/comments.xml")
    comments_extended_xml = docx_zip.read("word/commentsExtended.xml")
    document_xml = docx_zip.read("word/document.xml")
    with open("comments.xml", "wb") as f:
        f.write(comments_xml)
    with open("comments_ex.xml", "wb") as f:
        f.write(comments_extended_xml)
    with open("comments_of.xml", "wb") as f:
        f.write(document_xml)

    et_comments = ET.XML(comments_xml)
    et_comments_ex = ET.XML(comments_extended_xml)
    et_document = ET.XML(document_xml)
    comments = et_comments.xpath("//w:comment", namespaces=ooXMLns)
    comments_ex = et_comments_ex.xpath("//w15:commentEx", namespaces=ooXMLns)
    comment_doc_ranges = et_document.xpath("//w:commentRangeStart", namespaces=ooXMLns)

    comments_dict = {}
    comments_ex_dict = {}
    comments_doc_dict = {}
    parent_child_para_ids = {}

    # get information about relationships between comments and whether they are resolved
    for c in comments_ex:
        para_id = c.xpath("@w15:paraId", namespaces=ooXMLns)[0]
        parent_para_id_tmp = c.xpath("@w15:paraIdParent", namespaces=ooXMLns)
        if len(parent_para_id_tmp) != 0:
            reply = True
            parent_id = parent_para_id_tmp[0]
        else:
            reply = False
            parent_id = None
        done = True if c.xpath("@w15:done", namespaces=ooXMLns)[0] == "1" else False
        comments_ex_dict[para_id] = {"is_reply": reply, "parent_id": parent_id, "resolved": done}
        if parent_id is not None:
            if parent_id in parent_child_para_ids:
                parent_child_para_ids[parent_id].append(para_id)
            else:
                parent_child_para_ids[parent_id] = [para_id]

    for c in comments:
        comment = c.xpath("string(.)", namespaces=ooXMLns)
        comment_id = c.xpath("@w:id", namespaces=ooXMLns)[0]
        comment_author = c.xpath("@w:author", namespaces=ooXMLns)[0]
        comment_initials = c.xpath("@w:initials", namespaces=ooXMLns)[0]
        comment_date = c.xpath("@w:date", namespaces=ooXMLns)[0]
        comment_para_id = c.xpath("w:p/@w14:paraId", namespaces=ooXMLns)[0]
        comments_dict[comment_id] = {
            "para_id": comment_para_id,
            "author": comment_author,
            "initials": comment_initials,
            "date": comment_date,
            "comment": comment,
        }
    for c in comment_doc_ranges:
        comments_of_id = c.xpath("@w:id", namespaces=ooXMLns)[0]
        parts = c.xpath(
            "//w:r[preceding::w:commentRangeStart[@w:id="
            + comments_of_id
            + "] and following::w:commentRangeEnd[@w:id="
            + comments_of_id
            + "]]",
            namespaces=ooXMLns,
        )
        comment_of = ""
        for part in parts:
            # assumes each part is a new paragraph, but can't cope with entirely blank lines
            comment_of += part.xpath("string(.)", namespaces=ooXMLns) + "\n"
        comments_doc_dict[comments_of_id] = comment_of

    return parent_child_para_ids, comments_dict, comments_ex_dict, comments_doc_dict


if __name__ == "__main__":
    filename = r"C:\Users\hudsonm0098\Downloads\stuff to try out.docx"
    result = get_document_comments(filename)
    dummy = 1
