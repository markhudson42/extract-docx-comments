import zipfile
import xlwings as xw

from lxml import etree as ET
from dataclasses import dataclass


@dataclass
class DocxComments:
    """Class for storing information about docx Word document comments."""

    comments: dict
    comments_ex: dict
    comments_doc: dict


# XML mnamespaces for the tags we want
ooXMLns = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}


def get_document_comments(docx_fileName):
    print(f"Reading data from: {docx_fileName}")
    docx_zip = zipfile.ZipFile(docx_fileName)

    comments_file = docx_zip.read("word/comments.xml")
    comments_extended_file = docx_zip.read("word/commentsExtended.xml")
    document_file = docx_zip.read("word/document.xml")
    with open("comments.xml", "wb") as f:
        f.write(comments_file)
    with open("comments_ex.xml", "wb") as f:
        f.write(comments_extended_file)
    with open("comments_of.xml", "wb") as f:
        f.write(document_file)

    et_comments = ET.XML(comments_file)
    et_comments_ex = ET.XML(comments_extended_file)
    et_document = ET.XML(document_file)
    comments = et_comments.xpath("//w:comment", namespaces=ooXMLns)
    comments_ex = et_comments_ex.xpath("//w15:commentEx", namespaces=ooXMLns)
    comment_doc_ranges = et_document.xpath("//w:commentRangeStart", namespaces=ooXMLns)

    comments_dict = {}
    comments_ex_dict = {}
    comments_doc_dict = {}

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

    docx_comments = DocxComments(comments_dict, comments_ex_dict, comments_doc_dict)
    return docx_comments


def process_comment(comment_id, parent_id, comment_data, comments_doc):
    global comments_seen
    global parent_child_relationships
    global number_processed
    global sht

    if comment_id in comments_seen:
        return
    comments_seen.append(comment_id)
    number_processed += 1

    comment_author = comment_data["author"]
    comment_date = comment_data["date"]
    comment_para_id = comment_data["para_id"]
    is_reply = docx_comments.comments_ex[comment_para_id]["is_reply"]
    resolved = docx_comments.comments_ex[comment_para_id]["resolved"]
    reply_to = parent_id if is_reply else "n/a"
    comment_text = comment_data["comment"]
    doc_text = comments_doc[comment_id]

    output_data = [
        comment_id,
        "Yes" if resolved else "No",
        "Yes" if is_reply else "No",
        reply_to,
        comment_author,
        comment_date,
        doc_text,
        comment_text,
    ]
    sht.range(number_processed + 1, 1).value = output_data

    if comment_id in parent_child_relationships:
        child_comments = parent_child_relationships[comment_id]
        if len(child_comments) != 0:
            for reply_id in child_comments:
                comment_data = docx_comments.comments[reply_id]
                process_comment(reply_id, comment_id, comment_data, comments_doc)


if __name__ == "__main__":
    filename = r"stuff to try out.docx"
    output_workbook = "docx_comments_output.xlsx"

    docx_comments = get_document_comments(filename)

    # get the comment ID associated with each para ID
    para_id_to_comment_id = {comment["para_id"]: id for id, comment in docx_comments.comments.items()}

    # get the relationships between the parent and child comments
    parent_child_relationships = {}
    for comment_para_id, v in docx_comments.comments_ex.items():
        if v["is_reply"]:
            parent_para_id = v["parent_id"]
            parent_comment_id = para_id_to_comment_id[parent_para_id]
            child_comment_id = para_id_to_comment_id[comment_para_id]
            if parent_comment_id not in parent_child_relationships:
                parent_child_relationships[parent_comment_id] = [child_comment_id]
            else:
                parent_child_relationships[parent_comment_id].append(child_comment_id)

    # list the details of the different comments, grouping replies with their parents
    comments_seen = []
    number_processed = 0

    # TODO: create a spreadsheet with the comment data in
    with xw.App() as app:
        book = xw.Book()
        sheets = book.sheets
        if len(sheets) == 0:
            sht = book.sheets.add("Sheet 1")
        else:
            sht = sheets[0]

        sht.range(1, 1).value = ["ID", "Resolved?", "Reply?", "Reply To", "Author", "Date", "Doc Text", "Comment"]
        for comment_id, comment_data in docx_comments.comments.items():
            process_comment(comment_id, comment_id, comment_data, docx_comments.comments_doc)

        sht.range(1, 6).column_width = 60
        sht.range(1, 7).column_width = 60
        sht.autofit(axis="columns")
        sht.autofit(axis="rows")

        # save the workbook
        book.save(output_workbook)
        book.close()

    dummy = 1
