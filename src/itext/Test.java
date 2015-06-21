package itext;

import java.awt.Color;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.rtf.RtfWriter2;
import com.lowagie.text.rtf.style.RtfParagraphStyle;

public class Test {
    public static void main(String[] args) throws DocumentException, FileNotFoundException {
        Document document = new Document(PageSize.A4.rotate());
        RtfWriter2.getInstance(document, new FileOutputStream("d:\\test.doc"));
        document.open();
        Font titleFont = new Font(Font.NORMAL,16, Font.BOLD);
/* ���ñ���1��ʽ */
        RtfParagraphStyle rtfGsBt1 = RtfParagraphStyle.STYLE_HEADING_1;
        rtfGsBt1.setAlignment(Element.ALIGN_LEFT);
        rtfGsBt1.setStyle(Font.NORMAL);
        rtfGsBt1.setSize(13);
        rtfGsBt1.setColor(new Color(91,220,42));
/* ���ñ���2��ʽ */
        RtfParagraphStyle rtfGsBt2 = RtfParagraphStyle.STYLE_HEADING_2;
        rtfGsBt2.setAlignment(Element.ALIGN_LEFT);
        rtfGsBt2.setStyle(Font.NORMAL);
        rtfGsBt2.setSize(12);
        Paragraph title = new Paragraph("����");
        title.setAlignment(Element.ALIGN_CENTER);
        title.setFont(titleFont);
        document.add(title);
//����
        title = new Paragraph("1.��һ��");
        title.setFont(rtfGsBt1);
        document.add(title);
        title = new Paragraph("1.1 ��һ��");
        title.setFont(rtfGsBt2);
        document.add(title);
        title = new Paragraph("1.2 �ڶ���");
        title.setFont(rtfGsBt2);
        document.add(title);
        title = new Paragraph("2.�ڶ���");
        title.setFont(rtfGsBt1);
        document.add(title);
        title = new Paragraph("2.1 ��һ��");
        title.setFont(rtfGsBt2);
        document.add(title);
        title = new Paragraph("2.2 �ڶ���");
        title.setFont(rtfGsBt2);
        document.add(title);

        document.close();
    }

}
