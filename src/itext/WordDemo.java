package itext;

import java.awt.Color;
import java.io.IOException;

import com.lowagie.text.Cell;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Phrase;
import com.lowagie.text.Table;
import com.lowagie.text.rtf.style.RtfParagraphStyle;

/**
 * Description:
 * User: yangbin
 * Date: 15-06-21
 * Time: 8:15
 */
public class WordDemo {
    public static void main(String[] args) throws IOException, DocumentException {
        WordUtils wordUtils = new WordUtils();
        wordUtils.openDocument("D:\\wordDemo.doc");

        //��һ��������ʽ
        RtfParagraphStyle rtfGsBt1 = RtfParagraphStyle.STYLE_HEADING_1;
        rtfGsBt1.setAlignment(Element.ALIGN_LEFT);
        rtfGsBt1.setStyle(Font.BOLD);
        rtfGsBt1.setSize(15);

        //�ڶ���������ʽ
        RtfParagraphStyle rtfGsBt2 = RtfParagraphStyle.STYLE_HEADING_2;
        rtfGsBt2.setAlignment(Element.ALIGN_LEFT);
        rtfGsBt2.setStyle(Font.BOLD);
        rtfGsBt2.setSize(13);

        wordUtils.insertTitle("����IText��дWord����", 17, Font.BOLD, Element.ALIGN_CENTER);

        wordUtils.insertTitlePattern("һ����������",rtfGsBt1);
        wordUtils.insertTitlePatternSecond("1.1 ��һС������", rtfGsBt2);
        wordUtils.insertContext("iText�������Ŀ���Դ���վ��sourceforgeһ����Ŀ������������PDF�ĵ���һ��java��⡣" +
                "ͨ��iText������������PDF��rtf���ĵ������ҿ��Խ�XML��Html�ļ�ת��ΪPDF�ļ��� iText�İ�װ�ǳ����㣬" +
                "����iText.jar�ļ���ֻ��Ҫ��ϵͳ��CLASSPATH�м���iText.jar��·�����ڳ����оͿ���ʹ��iText����ˡ�", 12, Font.NORMAL, Element.ALIGN_CENTER);
        wordUtils.insertTitlePatternSecond("1.2 �ڶ�С������", rtfGsBt2);
        wordUtils.insertContext("һ������£�iTextʹ����������һ��Ҫ�����Ŀ�У�\n" +
                "    �����޷���ǰ���ã�ȡ�����û��������ʵʱ�����ݿ���Ϣ��\n" +
                "    �������ݣ�ҳ����࣬PDF�ĵ������ֶ����ɡ�\n" +
                "    �ĵ��������˲��룬������ģʽ���Զ�������\n" +
                "    ���ݱ����ƻ���Ի������磬�ն˿ͻ���������Ҫ����ڴ�����ҳ���ϡ�\n", 12, Font.NORMAL, Element.ALIGN_LEFT);

        wordUtils.insertTitlePattern("��������ͼƬ",rtfGsBt1);
        //Ϊ�˷�����ʹ�ã�����ͼƬ��·����ʹ������ͼƬ��·�����������Ҫ�����������޸�Ϊ����·����
        wordUtils.insertImg("http://sports.nen.com.cn/imagelist/11/24/w26o1rh8621m.jpg", Element.ALIGN_CENTER, 400, 400, 50, 50, 50, 0);

        wordUtils.insertTitlePattern("��������򵥱��",rtfGsBt1);
        wordUtils.getDocument().add(new Phrase(""));//Ϊ�˽���ڲ���������֮���������Ŀո�
        new Paragraph();
        wordUtils.insertSimpleTable(4, 3);

        wordUtils.insertTitlePattern("�ġ�����ϸ��ӱ��",rtfGsBt1);
        wordUtils.getDocument().add(new Paragraph("22"));
        wordUtils.getDocument().add(new WordDemo().insertBGColor());
        wordUtils.getDocument().add(new WordDemo().insertComplexTable());

        wordUtils.closeDocument();

    }

    /**
     * @return ���б�����ɫ��table
     * @throws DocumentException
     */
    public Table insertBGColor() throws DocumentException {
        Table table = new Table(4);//����һ�����еı��
        int width[] = {25, 25, 25, 25};
        table.setWidths(width);//����ϵ����ռ����
        table.setWidth(100);
        table.setAutoFillEmptyCells(true);
        table.setAlignment(Element.ALIGN_CENTER);//������ʾ
        table.setAlignment(Element.ALIGN_MIDDLE);//��ֱ������ʾ
        table.setBorder(30);
        table.setBorderWidth(230);//�߿���
        table.setSpacing(2);
        table.setPadding(3);
        table.setBorderColor(new Color(58, 255, 132));//�߿���ɫ

        Cell cell = new Cell(new Phrase("��һ"));
        cell.setVerticalAlignment(Element.ALIGN_CENTER);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setBorderColor(new Color(189, 22, 33));
        cell.setBackgroundColor(new Color(58, 137, 20));
        table.addCell(cell);

        Cell cell2 = new Cell(new Phrase("�ж�"));
        cell2.setVerticalAlignment(Element.ALIGN_CENTER);
        cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell2.setBorderColor(new Color(189, 22, 33));
        cell2.setBackgroundColor(new Color(137, 34, 44));
        table.addCell(cell2);

        Cell cell3 = new Cell(new Phrase("����"));
        cell3.setVerticalAlignment(Element.ALIGN_CENTER);
        cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell3.setBorderColor(new Color(189, 22, 33));
        cell3.setBackgroundColor(new Color(232, 219, 48));
        table.addCell(cell3);

        Cell cell4 = new Cell(new Phrase("����"));
        cell4.setVerticalAlignment(Element.ALIGN_CENTER);
        cell4.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell4.setBorderColor(new Color(189, 22, 33));
        cell4.setBackgroundColor(new Color(54, 246, 231));
        table.addCell(cell);

        for (int i = 0; i < 8; i++) {
            table.addCell(new Cell("�Զ�������"));
        }
        return table;
    }

    /**
     * @return ���ϱ��ļ�����
     * @throws DocumentException
     */
    public Table insertComplexTable() throws DocumentException {
        Table table = new Table(5);
        int width[] = {20, 20, 20, 20, 20};
        table.setWidths(width);//����ϵ����ռ����
        table.setWidth(100);
        table.setAutoFillEmptyCells(true);
        table.setAlignment(Element.ALIGN_CENTER);//������ʾ
        table.setAlignment(Element.ALIGN_MIDDLE);//��ֱ������ʾ
        table.setBorder(1000);
        table.setBorderWidth(1);//�߿���
        //table.setSpacing(2);
        //table.setPadding(3);
        table.setBorderColor(new Color(58, 255, 132));//�߿���ɫ

        Cell cell = new Cell(new Phrase("ռ�����еĵ�Ԫ��"));
        cell.setColspan(3);//���õ�ǰ��Ԫ��ռ�ݵ�����
        cell.setVerticalAlignment(Element.ALIGN_CENTER);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setBorderColor(new Color(189, 22, 33));
        cell.setBackgroundColor(new Color(58, 137, 20));
        table.addCell(cell);

        Cell cell2 = new Cell(new Phrase("������"));
        cell2.setVerticalAlignment(Element.ALIGN_CENTER);
        cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell2.setBorderColor(new Color(189, 22, 33));
        // cell2.setBackgroundColor(new Color(137, 34, 44));
        table.addCell(cell2);

        Cell cell3 = new Cell(new Phrase("������"));
        cell3.setVerticalAlignment(Element.ALIGN_CENTER);
        cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell3.setBorderColor(new Color(189, 22, 33));
        // cell3.setBackgroundColor(new Color(232, 219, 48));
        table.addCell(cell3);

        Cell cell4 = new Cell(new Phrase("ռ�����еĵ�Ԫ��"));
        cell4.setBackgroundColor(new Color(232, 219, 48));
        cell4.setRowspan(2);
        table.addCell(cell4);

        for (int i = 0; i < 8; i++) {
            table.addCell(new Cell("�Զ�������"));
        }

        Cell cell5 = new Cell(new Phrase("ռ���������еĵ�Ԫ��"));
        cell5.setBackgroundColor(new Color(137, 34, 44));
        cell5.setRowspan(2);
        cell5.setColspan(2);
        table.addCell(cell5);

        for (int i = 0; i < 6; i++) {
            table.addCell(new Cell("�Զ�������"));
        }
        return table;
    }
}

