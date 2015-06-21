package itext;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;

import com.lowagie.text.Cell;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Table;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.rtf.RtfWriter2;
import com.lowagie.text.rtf.style.RtfParagraphStyle;

/**
 * word帮助文件
 */
public class WordUtils {
    private Document document;
    private BaseFont bfChinese;
    private String test;

    public BaseFont getBfChinese() {
        return bfChinese;  
    }  
  
    public void setBfChinese(BaseFont bfChinese) {  
        this.bfChinese = bfChinese;  
    }  
  
    public Document getDocument() {  
        return document;  
    }  
  
    public void setDocument(Document document) {  
        this.document = document;  
    }  
  
    public WordUtils(){  
        this.document = new Document(PageSize.A4);//设置纸张大小
          
    }  
    /** 建立一个书写器(Writer)与document对象关联，通过书写器(Writer)可以将文档写入到磁盘中
     * @param filePath  要操作的文档路径，若文档不存在会自动创建
     * @throws com.lowagie.text.DocumentException
     * @throws java.io.IOException
     */  
    public void openDocument(String filePath) throws DocumentException,  
    IOException {  
//       建立一个书写器(Writer)与document对象关联，通过书写器(Writer)可以将文档写入到磁盘中  
        RtfWriter2.getInstance(this.document, new FileOutputStream(filePath));  
        this.document.open();  
//       设置中文字体  
        this.bfChinese = BaseFont.createFont("STSongStd-Light",  
                "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
    }
      
    /** 
     * @param titleStr 标题
     * @param fontsize 字体大小 
     * @param fontStyle 字体样式 
     * @param elementAlign 对齐方式 
     * @throws com.lowagie.text.DocumentException
     */  
    public void insertTitle(String titleStr,int fontsize,int fontStyle,int elementAlign) throws DocumentException{  
        Font titleFont = new Font(this.bfChinese, fontsize, fontStyle);  
        Paragraph title = new Paragraph(titleStr);  
        // 设置标题格式对齐方式  
        title.setAlignment(elementAlign);  
        title.setFont(titleFont);  
          
        this.document.add(title);  
    }  
      


    /**
     * 设置带有目录格式的标题(标题1格式)
     *
     * @param rtfParagraphStyle 标题1样式
     * @param titleStr 标题
     * @throws DocumentException
     */
    public void insertTitlePattern(String titleStr, RtfParagraphStyle rtfParagraphStyle) throws DocumentException{
        Paragraph title = new Paragraph(titleStr);
        title.setFont(rtfParagraphStyle);
        this.document.add(title);
    }  
    

    /**
     * 设置带有目录格式的标题(标题2格式)
     * @param titleStr 标题
     * @param rtfParagraphStyle 标题2样式
     * @throws DocumentException
     */
    public void insertTitlePatternSecond(String titleStr,RtfParagraphStyle rtfParagraphStyle) throws DocumentException{
        Paragraph title = new Paragraph(titleStr);
        // 设置标题格式对齐方式  
        title.setFont(rtfParagraphStyle);
        this.document.add(title);
    }  
    
    
    /** 
     * @param tableName 标题 
     * @param fontsize 字体大小 
     * @param fontStyle 字体样式 
     * @param elementAlign 对齐方式 
     * @throws com.lowagie.text.DocumentException
     */  
    public void insertTableName(String tableName,int fontsize,int fontStyle,int elementAlign) throws DocumentException{  
        Font titleFont = new Font(this.bfChinese, fontsize, fontStyle);  
        titleFont.setColor(102, 102, 153);
        Paragraph title = new Paragraph(tableName);  
        // 设置标题格式对齐方式  
        title.setAlignment(elementAlign);  
        title.setFont(titleFont);  
          
        this.document.add(title);  
    }  
    
    /** 
     * @param contextStr 内容 
     * @param fontsize 字体大小 
     * @param fontStyle 字体样式 
     * @param elementAlign 对齐方式 
     * @throws com.lowagie.text.DocumentException
     */  
    public void insertContext(String contextStr,int fontsize,int fontStyle,int elementAlign) throws DocumentException{  
        // 正文字体风格  
        Font contextFont = new Font(bfChinese, fontsize, fontStyle);  
        Paragraph context = new Paragraph(contextStr);  
        //设置行距  
        context.setLeading(3f);
        // 正文格式左对齐  
      //  context.setAlignment(elementAlign);
        context.setFont(contextFont);  
        // 离上一段落（标题）空的行数  
        context.setSpacingBefore(1);
        // 设置第一行空的列数  
        context.setFirstLineIndent(20);  
        document.add(context);  
    }


    /**
     * @param imgUrl 图片路径
     * @param imageAlign 显示位置
     * @param height 显示高度
     * @param weight 显示宽度
     * @param percent 显示比例
     * @param heightPercent 显示高度比例
     * @param weightPercent 显示宽度比例
     * @param rotation 显示图片旋转角度
     * @throws java.net.MalformedURLException
     * @throws java.io.IOException
     * @throws com.lowagie.text.DocumentException
     */
    public void insertImg(String imgUrl,int imageAlign,int height,int weight,int percent,int heightPercent,int weightPercent,int rotation) throws MalformedURLException, IOException, DocumentException{  
//       添加图片  
        Image img = Image.getInstance(imgUrl);  
        if(img==null)  
            return;  
        img.setAbsolutePosition(0, 0);  
        img.setAlignment(imageAlign);  
        img.scaleAbsolute(height, weight);
        img.scaleAbsolute(1000, 1000);
        img.scalePercent(percent);
        img.scalePercent(heightPercent, weightPercent);  
        img.setRotation(rotation);  
        document.add(img);
    }

    /**
     * 添加简单表格
     * @param column 表格列数(必须)
     * @param row 表格行数
     * @throws DocumentException
     */
    public void insertSimpleTable(int column,int row) throws DocumentException {
        Table table=new Table(column);//列数必须设置，而行数则可以按照个人要求来决定是否需要设置
        table.setAlignment(Element.ALIGN_CENTER);// 居中显示
        table.setAlignment(Element.ALIGN_MIDDLE);// 纵向居中显示
        table.setAutoFillEmptyCells(true);// 自动填满
        table.setBorderColor(new Color(0, 125, 255));// 边框颜色
        table.setBorderWidth(1);// 边框宽度
        table.setSpacing(2);// 衬距，
        table.setPadding(2);// 即单元格之间的间距
        table.setBorder(20);// 边框
        for (int i = 0; i < column*3; i++) {
            table.addCell(new Cell(""+i));
        }
        document.add(table);
    }
    /**
     * 在操作完成后必须关闭document,否则即使生成了word文档，打开也会发生错误
     * @throws DocumentException
     */
    public void closeDocument() throws DocumentException{  
        this.document.close();  
    }  
    
}

