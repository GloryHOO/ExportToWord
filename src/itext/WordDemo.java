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

        //第一级标题样式
        RtfParagraphStyle rtfGsBt1 = RtfParagraphStyle.STYLE_HEADING_1;
        rtfGsBt1.setAlignment(Element.ALIGN_LEFT);
        rtfGsBt1.setStyle(Font.BOLD);
        rtfGsBt1.setSize(15);

        //第二级标题样式
        RtfParagraphStyle rtfGsBt2 = RtfParagraphStyle.STYLE_HEADING_2;
        rtfGsBt2.setAlignment(Element.ALIGN_LEFT);
        rtfGsBt2.setStyle(Font.BOLD);
        rtfGsBt2.setSize(13);

        wordUtils.insertTitle("利用IText书写Word例子", 17, Font.BOLD, Element.ALIGN_CENTER);

        wordUtils.insertTitlePattern("一、基本内容",rtfGsBt1);
        wordUtils.insertTitlePatternSecond("1.1 第一小节内容", rtfGsBt2);
        wordUtils.insertContext("iText是著名的开放源码的站点sourceforge一个项目，是用于生成PDF文档的一个java类库。" +
                "通过iText不仅可以生成PDF或rtf的文档，而且可以将XML、Html文件转化为PDF文件。 iText的安装非常方便，" +
                "下载iText.jar文件后，只需要在系统的CLASSPATH中加入iText.jar的路径，在程序中就可以使用iText类库了。", 12, Font.NORMAL, Element.ALIGN_CENTER);
        wordUtils.insertTitlePatternSecond("1.2 第二小节内容", rtfGsBt2);
        wordUtils.insertContext("一般情况下，iText使用在有以下一个要求的项目中：\n" +
                "    内容无法提前利用：取决于用户的输入或实时的数据库信息。\n" +
                "    由于内容，页面过多，PDF文档不能手动生成。\n" +
                "    文档需在无人参与，批处理模式下自动创建。\n" +
                "    内容被定制或个性化；例如，终端客户的名字需要标记在大量的页面上。\n", 12, Font.NORMAL, Element.ALIGN_LEFT);

        wordUtils.insertTitlePattern("二、插入图片",rtfGsBt1);
        //为了方便大家使用，下面图片的路径是使用网上图片的路径，如果有需要，可以自行修改为本地路径。
        wordUtils.insertImg("http://sports.nen.com.cn/imagelist/11/24/w26o1rh8621m.jpg", Element.ALIGN_CENTER, 400, 400, 50, 50, 50, 0);

        wordUtils.insertTitlePattern("三、插入简单表格",rtfGsBt1);
        wordUtils.getDocument().add(new Phrase(""));//为了解决在插入段落与表之间产生多余的空格
        new Paragraph();
        wordUtils.insertSimpleTable(4, 3);

        wordUtils.insertTitlePattern("四、插入较复杂表格",rtfGsBt1);
        wordUtils.getDocument().add(new Paragraph("22"));
        wordUtils.getDocument().add(new WordDemo().insertBGColor());
        wordUtils.getDocument().add(new WordDemo().insertComplexTable());

        wordUtils.closeDocument();

    }

    /**
     * @return 带有背景颜色的table
     * @throws DocumentException
     */
    public Table insertBGColor() throws DocumentException {
        Table table = new Table(4);//生成一个四列的表格
        int width[] = {25, 25, 25, 25};
        table.setWidths(width);//设置系列所占比例
        table.setWidth(100);
        table.setAutoFillEmptyCells(true);
        table.setAlignment(Element.ALIGN_CENTER);//居中显示
        table.setAlignment(Element.ALIGN_MIDDLE);//垂直居中显示
        table.setBorder(30);
        table.setBorderWidth(230);//边框宽度
        table.setSpacing(2);
        table.setPadding(3);
        table.setBorderColor(new Color(58, 255, 132));//边框颜色

        Cell cell = new Cell(new Phrase("列一"));
        cell.setVerticalAlignment(Element.ALIGN_CENTER);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setBorderColor(new Color(189, 22, 33));
        cell.setBackgroundColor(new Color(58, 137, 20));
        table.addCell(cell);

        Cell cell2 = new Cell(new Phrase("列二"));
        cell2.setVerticalAlignment(Element.ALIGN_CENTER);
        cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell2.setBorderColor(new Color(189, 22, 33));
        cell2.setBackgroundColor(new Color(137, 34, 44));
        table.addCell(cell2);

        Cell cell3 = new Cell(new Phrase("列三"));
        cell3.setVerticalAlignment(Element.ALIGN_CENTER);
        cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell3.setBorderColor(new Color(189, 22, 33));
        cell3.setBackgroundColor(new Color(232, 219, 48));
        table.addCell(cell3);

        Cell cell4 = new Cell(new Phrase("列四"));
        cell4.setVerticalAlignment(Element.ALIGN_CENTER);
        cell4.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell4.setBorderColor(new Color(189, 22, 33));
        cell4.setBackgroundColor(new Color(54, 246, 231));
        table.addCell(cell);

        for (int i = 0; i < 8; i++) {
            table.addCell(new Cell("自定义内容"));
        }
        return table;
    }

    /**
     * @return 复合表格的简单例子
     * @throws DocumentException
     */
    public Table insertComplexTable() throws DocumentException {
        Table table = new Table(5);
        int width[] = {20, 20, 20, 20, 20};
        table.setWidths(width);//设置系列所占比例
        table.setWidth(100);
        table.setAutoFillEmptyCells(true);
        table.setAlignment(Element.ALIGN_CENTER);//居中显示
        table.setAlignment(Element.ALIGN_MIDDLE);//垂直居中显示
        table.setBorder(1000);
        table.setBorderWidth(1);//边框宽度
        //table.setSpacing(2);
        //table.setPadding(3);
        table.setBorderColor(new Color(58, 255, 132));//边框颜色

        Cell cell = new Cell(new Phrase("占据三列的单元格"));
        cell.setColspan(3);//设置当前单元格占据的列数
        cell.setVerticalAlignment(Element.ALIGN_CENTER);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setBorderColor(new Color(189, 22, 33));
        cell.setBackgroundColor(new Color(58, 137, 20));
        table.addCell(cell);

        Cell cell2 = new Cell(new Phrase("第四列"));
        cell2.setVerticalAlignment(Element.ALIGN_CENTER);
        cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell2.setBorderColor(new Color(189, 22, 33));
        // cell2.setBackgroundColor(new Color(137, 34, 44));
        table.addCell(cell2);

        Cell cell3 = new Cell(new Phrase("第五列"));
        cell3.setVerticalAlignment(Element.ALIGN_CENTER);
        cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell3.setBorderColor(new Color(189, 22, 33));
        // cell3.setBackgroundColor(new Color(232, 219, 48));
        table.addCell(cell3);

        Cell cell4 = new Cell(new Phrase("占据两行的单元格"));
        cell4.setBackgroundColor(new Color(232, 219, 48));
        cell4.setRowspan(2);
        table.addCell(cell4);

        for (int i = 0; i < 8; i++) {
            table.addCell(new Cell("自定义内容"));
        }

        Cell cell5 = new Cell(new Phrase("占据两行两列的单元格"));
        cell5.setBackgroundColor(new Color(137, 34, 44));
        cell5.setRowspan(2);
        cell5.setColspan(2);
        table.addCell(cell5);

        for (int i = 0; i < 6; i++) {
            table.addCell(new Cell("自定义内容"));
        }
        return table;
    }
}

