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
 * word�����ļ�
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
        this.document = new Document(PageSize.A4);//����ֽ�Ŵ�С
          
    }  
    /** ����һ����д��(Writer)��document���������ͨ����д��(Writer)���Խ��ĵ�д�뵽������
     * @param filePath  Ҫ�������ĵ�·�������ĵ������ڻ��Զ�����
     * @throws com.lowagie.text.DocumentException
     * @throws java.io.IOException
     */  
    public void openDocument(String filePath) throws DocumentException,  
    IOException {  
//       ����һ����д��(Writer)��document���������ͨ����д��(Writer)���Խ��ĵ�д�뵽������  
        RtfWriter2.getInstance(this.document, new FileOutputStream(filePath));  
        this.document.open();  
//       ������������  
        this.bfChinese = BaseFont.createFont("STSongStd-Light",  
                "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
    }
      
    /** 
     * @param titleStr ����
     * @param fontsize �����С 
     * @param fontStyle ������ʽ 
     * @param elementAlign ���뷽ʽ 
     * @throws com.lowagie.text.DocumentException
     */  
    public void insertTitle(String titleStr,int fontsize,int fontStyle,int elementAlign) throws DocumentException{  
        Font titleFont = new Font(this.bfChinese, fontsize, fontStyle);  
        Paragraph title = new Paragraph(titleStr);  
        // ���ñ����ʽ���뷽ʽ  
        title.setAlignment(elementAlign);  
        title.setFont(titleFont);  
          
        this.document.add(title);  
    }  
      


    /**
     * ���ô���Ŀ¼��ʽ�ı���(����1��ʽ)
     *
     * @param rtfParagraphStyle ����1��ʽ
     * @param titleStr ����
     * @throws DocumentException
     */
    public void insertTitlePattern(String titleStr, RtfParagraphStyle rtfParagraphStyle) throws DocumentException{
        Paragraph title = new Paragraph(titleStr);
        title.setFont(rtfParagraphStyle);
        this.document.add(title);
    }  
    

    /**
     * ���ô���Ŀ¼��ʽ�ı���(����2��ʽ)
     * @param titleStr ����
     * @param rtfParagraphStyle ����2��ʽ
     * @throws DocumentException
     */
    public void insertTitlePatternSecond(String titleStr,RtfParagraphStyle rtfParagraphStyle) throws DocumentException{
        Paragraph title = new Paragraph(titleStr);
        // ���ñ����ʽ���뷽ʽ  
        title.setFont(rtfParagraphStyle);
        this.document.add(title);
    }  
    
    
    /** 
     * @param tableName ���� 
     * @param fontsize �����С 
     * @param fontStyle ������ʽ 
     * @param elementAlign ���뷽ʽ 
     * @throws com.lowagie.text.DocumentException
     */  
    public void insertTableName(String tableName,int fontsize,int fontStyle,int elementAlign) throws DocumentException{  
        Font titleFont = new Font(this.bfChinese, fontsize, fontStyle);  
        titleFont.setColor(102, 102, 153);
        Paragraph title = new Paragraph(tableName);  
        // ���ñ����ʽ���뷽ʽ  
        title.setAlignment(elementAlign);  
        title.setFont(titleFont);  
          
        this.document.add(title);  
    }  
    
    /** 
     * @param contextStr ���� 
     * @param fontsize �����С 
     * @param fontStyle ������ʽ 
     * @param elementAlign ���뷽ʽ 
     * @throws com.lowagie.text.DocumentException
     */  
    public void insertContext(String contextStr,int fontsize,int fontStyle,int elementAlign) throws DocumentException{  
        // ����������  
        Font contextFont = new Font(bfChinese, fontsize, fontStyle);  
        Paragraph context = new Paragraph(contextStr);  
        //�����о�  
        context.setLeading(3f);
        // ���ĸ�ʽ�����  
      //  context.setAlignment(elementAlign);
        context.setFont(contextFont);  
        // ����һ���䣨���⣩�յ�����  
        context.setSpacingBefore(1);
        // ���õ�һ�пյ�����  
        context.setFirstLineIndent(20);  
        document.add(context);  
    }


    /**
     * @param imgUrl ͼƬ·��
     * @param imageAlign ��ʾλ��
     * @param height ��ʾ�߶�
     * @param weight ��ʾ���
     * @param percent ��ʾ����
     * @param heightPercent ��ʾ�߶ȱ���
     * @param weightPercent ��ʾ��ȱ���
     * @param rotation ��ʾͼƬ��ת�Ƕ�
     * @throws java.net.MalformedURLException
     * @throws java.io.IOException
     * @throws com.lowagie.text.DocumentException
     */
    public void insertImg(String imgUrl,int imageAlign,int height,int weight,int percent,int heightPercent,int weightPercent,int rotation) throws MalformedURLException, IOException, DocumentException{  
//       ���ͼƬ  
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
     * ��Ӽ򵥱��
     * @param column �������(����)
     * @param row �������
     * @throws DocumentException
     */
    public void insertSimpleTable(int column,int row) throws DocumentException {
        Table table=new Table(column);//�����������ã�����������԰��ո���Ҫ���������Ƿ���Ҫ����
        table.setAlignment(Element.ALIGN_CENTER);// ������ʾ
        table.setAlignment(Element.ALIGN_MIDDLE);// ���������ʾ
        table.setAutoFillEmptyCells(true);// �Զ�����
        table.setBorderColor(new Color(0, 125, 255));// �߿���ɫ
        table.setBorderWidth(1);// �߿���
        table.setSpacing(2);// �ľ࣬
        table.setPadding(2);// ����Ԫ��֮��ļ��
        table.setBorder(20);// �߿�
        for (int i = 0; i < column*3; i++) {
            table.addCell(new Cell(""+i));
        }
        document.add(table);
    }
    /**
     * �ڲ�����ɺ����ر�document,����ʹ������word�ĵ�����Ҳ�ᷢ������
     * @throws DocumentException
     */
    public void closeDocument() throws DocumentException{  
        this.document.close();  
    }  
    
}

