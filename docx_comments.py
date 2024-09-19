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


def get_author_and_date(comment, ns):
    author_tmp = comment.xpath("@w:author", namespaces=ns)
    initials_tmp = comment.xpath("@w:initials", namespaces=ns)
    date_tmp = comment.xpath("@w:date", namespaces=ns)

    comment_author = author_tmp[0] if len(author_tmp) > 0 else "MISSING"
    comment_initials = initials_tmp[0] if len(initials_tmp) > 0 else "MISSING"
    comment_date = date_tmp[0] if len(date_tmp) > 0 else "MISSING"

    return (comment_author, comment_initials, comment_date)


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
    with open("comments_doc.xml", "wb") as f:
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

    # read the comments themselves and their attributes, such as author, date and the paragraph ID
    for c in comments:
        comment = c.xpath("string(.)", namespaces=ooXMLns).strip()
        comment_id = c.xpath("@w:id", namespaces=ooXMLns)[0]
        comment_para_ids = c.xpath("w:p/@w14:paraId", namespaces=ooXMLns)  # there can be more than one ID per comment
        comment_author, comment_initials, comment_date = get_author_and_date(c, ooXMLns)
        comments_dict[comment_id] = {
            "para_ids": comment_para_ids,
            "author": comment_author,
            "initials": comment_initials,
            "date": comment_date,
            "comment": comment,
        }

    # read the document data and extract the text between comment ranges
    for c in comment_doc_ranges:
        comment_doc_id = c.xpath("@w:id", namespaces=ooXMLns)[0]
        # a comment can span multiple paragraphs, so we find w:r tags that
        # have a preceding sibling commentRangeStart tag with the same ID
        # as the comment - because the <w:r> is always after the commentRangeStart
        # and is at the same nested level as that tag - and a following
        # commentRangeEnd tag anywhere, not necessarily a sibling, with the same ID
        parts = c.xpath(
            "//w:r[preceding-sibling::w:commentRangeStart[@w:id="
            + comment_doc_id
            + "] and following::w:commentRangeEnd[@w:id="
            + comment_doc_id
            + "]]",
            namespaces=ooXMLns,
        )
        doc_text = ""
        for part in parts:
            # TODO: need to be able to handle new paragraphs within the document text that is commented
            doc_text += part.xpath("string(.)", namespaces=ooXMLns)
        comments_doc_dict[comment_doc_id] = doc_text.strip()

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
    comment_para_ids = comment_data["para_ids"]
    is_reply = False
    resolved = False
    comment_ex_dict = docx_comments.comments_ex
    for comment_para_id in comment_para_ids:
        if comment_para_id not in comment_ex_dict:
            continue
        para_id_attribs = comment_ex_dict[comment_para_id]
        if para_id_attribs["is_reply"]:
            is_reply = True
        if para_id_attribs["resolved"]:
            resolved = True
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

    # add the comment data to the workbook
    sht.range(number_processed + 1, 1).value = output_data

    # process any replies to this top-level comment so they appear in the workbook in sequence
    if comment_id in parent_child_relationships:
        child_comments = parent_child_relationships[comment_id]
        if len(child_comments) != 0:
            for reply_id in child_comments:
                comment_data = docx_comments.comments[reply_id]
                process_comment(reply_id, comment_id, comment_data, comments_doc)


if __name__ == "__main__":
    # filename = r"Document for Markup Testing-AJ.docx"
    filename = r"Zone Review.docx"
    output_workbook = "docx_comments_output.xlsx"

    docx_comments = get_document_comments(filename)

    # get the comment ID associated with each para ID
    para_id_to_comment_id = {
        para_id: comment_id
        for comment_id, comment_data in docx_comments.comments.items()
        for para_id in comment_data["para_ids"]
    }

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

    # create a spreadsheet with the comment data in
    with xw.App(visible=True, add_book=False) as app:
        app.screen_updating = False
        book = xw.Book()
        sheets = book.sheets
        if len(sheets) == 0:
            sht = book.sheets.add("Sheet 1")
        else:
            sht = sheets[0]

        sht.range(1, 1).value = [
            "ID",
            "Is Resolved?",
            "Is a Reply?",
            "Reply To",
            "Author",
            "Date",
            "Doc Text",
            "Comment",
        ]
        for comment_id, comment_data in docx_comments.comments.items():
            process_comment(comment_id, comment_id, comment_data, docx_comments.comments_doc)

        sht.autofit(axis="columns")
        sht.range(1, 7).column_width = 65  # document text
        sht.range(1, 8).column_width = 65  # comment text
        sht.autofit(axis="rows")

        sht.range("A1:H1").api.Font.Bold = True
        sht.used_range.api.WrapText = True

        # save the workbook
        book.save(output_workbook)
        book.close()

    dummy = 1
