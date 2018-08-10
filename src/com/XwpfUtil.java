package com;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 根据模板导出word文件工具类
 *
 * @author lmb
 * @date 2017-3-14
 */
public class XwpfUtil
{


    /**
     * 导出word文件
     */
    public void exportWord(Map<String, Object> params, InputStream is, XwpfUtil xwpfUtil)
    {

        try
        {
            XWPFDocument doc = new XWPFDocument(is);

            xwpfUtil.replaceInPara(doc, params);
            xwpfUtil.replaceInTable(doc, params);
            FileOutputStream fileOutputStream = new FileOutputStream(new File("1.docx"));
            //把构造好的文档写入输出流
            doc.write(fileOutputStream);
            //关闭流
            xwpfUtil.close(fileOutputStream);
            xwpfUtil.close(is);
            fileOutputStream.flush();
            fileOutputStream.close();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
    }

    /**
     * 替换word模板文档段落中的变量
     *
     * @param doc    要替换的文档
     * @param params 参数
     */
    public void replaceInPara(XWPFDocument doc, Map<String, Object> params)
    {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        String key = null;
        LinkedList<XWPFParagraph> tmp = null;
        while (iterator.hasNext())
        {
            para = iterator.next();
            String paragraphText = para.getParagraphText();
            // 识别list段落
            if(paragraphText.startsWith("<%="))
            {
                key = paragraphText.substring(3);
                tmp = new LinkedList<XWPFParagraph>();
            }
            if(key != null){
                tmp.add(para);
            }
            if(paragraphText.startsWith("%>")){

                XWPFParagraph start = tmp.removeFirst();


                XWPFParagraph last = tmp.removeLast();
                int lastPos = doc.getPosOfParagraph(last);
                doc.removeBodyElement(lastPos);

                Object o = params.get(key);
                List<HashMap<String,Object>> list = (List<HashMap<String, Object>>) o;
                int size = list.size();

                if(size == 0)
                {
                    for (XWPFParagraph xwpfParagraph : tmp)
                    {
                        int pos = doc.getPosOfParagraph(xwpfParagraph);
                        doc.removeBodyElement(pos);
                    }
                } else {
                    for (int i = 0; i < size; i++) {
                        for (XWPFParagraph xwpfParagraph : tmp) {

                            XmlCursor xmlCursor = start.getCTP().newCursor();
                            XWPFParagraph newParagraph = doc.insertNewParagraph(xmlCursor);


                            cloneParagraph(newParagraph, xwpfParagraph);
                            this.replaceInPara(newParagraph, list.get(i));

//                            copyAllRunsToAnotherParagraph( xwpfParagraph,  );
                        }
                    }
                }

                int posOfParagraph = doc.getPosOfParagraph(start);
                doc.removeBodyElement(posOfParagraph);

                break;
            }
        }
        this.replaceInPara2(doc,params);
    }

    public void replaceInPara2(XWPFDocument doc, Map<String, Object> params)
    {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        while (iterator.hasNext()){
            XWPFParagraph para = iterator.next();
            this.replaceInPara(para, params);
        }
    }

    public static void cloneParagraph(XWPFParagraph clone, XWPFParagraph source) {
        CTPPr pPr = clone.getCTP().isSetPPr() ? clone.getCTP().getPPr() : clone.getCTP().addNewPPr();
        pPr.set(source.getCTP().getPPr());
        for (XWPFRun r : source.getRuns()) {
            XWPFRun nr = clone.createRun();
            cloneRun(nr, r);
        }
    }

    public static void cloneRun(XWPFRun clone, XWPFRun source) {
        CTRPr rPr = clone.getCTR().isSetRPr() ? clone.getCTR().getRPr() : clone.getCTR().addNewRPr();
        rPr.set(source.getCTR().getRPr());
        clone.setText(source.getText(0));
    }


    // Copy all runs from one paragraph to another, keeping the style unchanged
    private static void copyAllRunsToAnotherParagraph(XWPFParagraph oldPar, XWPFParagraph newPar) {
        final int DEFAULT_FONT_SIZE = 10;

        for (XWPFRun run : oldPar.getRuns()) {
            String textInRun = run.getText(0);
            if (StringUtils.isEmpty(textInRun)) {
                continue;
            }

            int fontSize = run.getFontSize();
            System.out.println("run text = '" + textInRun + "' , fontSize = " + fontSize);

            XWPFRun newRun = newPar.createRun();

            // Copy text
            newRun.setText(textInRun);

            // Apply the same style
            newRun.setFontSize( ( fontSize == -1) ? DEFAULT_FONT_SIZE : run.getFontSize() );
            newRun.setFontFamily( run.getFontFamily() );
            newRun.setBold( run.isBold() );
            newRun.setItalic( run.isItalic() );
            newRun.setStrike( run.isStrike() );
            newRun.setColor( run.getColor() );
        }
    }


    /**
     * 替换段落中的变量
     *
     * @param para   要替换的段落
     * @param params 替换参数
     */
    public void replaceInPara(XWPFParagraph para, Map<String, Object> params)
    {
        List<XWPFRun> runs;
        boolean get = false;
        String paragraphText = para.getParagraphText();


        if (((Matcher) this.matcher(paragraphText)).find())
        {
            runs = para.getRuns();
            int start = -1;
            int end = -1;
            String str = "";
            for (int i = 0; i < runs.size(); i++)
            {
                XWPFRun run = runs.get(i);

                String runText = run.toString().trim();

                if (StringUtils.isNotBlank(runText) && '$' == runText.charAt(0) && '{' == runText.charAt(1))
                {
                    start = i;
                }
                if (StringUtils.isNotBlank(runText) && (start != -1))
                {
                    str += runText;
                }
                if (StringUtils.isNotBlank(runText) && '}' == runText.charAt(runText.length() - 1))
                {
                    if (start != -1)
                    {
                        end = i;
                        get = true;
                        break;
                    }
                }
            }
            String replace = String.valueOf(params.get(str));

//            for (int i = start; i <= end; i++)
//            {
//                para.removeRun(i);
//                i--;
//                end--;
//                System.out.println("remove i=" + i);
//            }

            runs.get(start).setText(replace,0);
            int i = start;
            while(i < end){
                para.removeRun(start+1);
                i++;
            }
//            if (StringUtils.isBlank(str))
//            {
//                String temp = para.getParagraphText();
//                str = temp.trim().substring(temp.indexOf("${"), temp.indexOf("}") + 1);
//            }

        }
        if(get){
            replaceInPara( para, params);
        }
    }

    /**
     * 替换word模板文档表格中的变量
     *
     * @param doc    要替换的文档
     * @param params 参数
     */
    public void replaceInTable(XWPFDocument doc, Map<String, Object> params)
    {
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        while (iterator.hasNext())
        {
            table = iterator.next();
            rows = table.getRows();
            for (XWPFTableRow row : rows)
            {
                cells = row.getTableCells();
                for (XWPFTableCell cell : cells)
                {
                    paras = cell.getParagraphs();
                    for (XWPFParagraph para : paras)
                    {
                        this.replaceInPara(para, params);
                    }
                }
            }
        }
    }

    /**
     * 正则匹配字符串
     *
     * @return
     */
    public Object matcher(String str)
    {
        System.out.println("段落："+str);
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        return pattern.matcher(str);
    }

    /**
     * 关闭输入流
     *
     * @param is
     */
    public void close(InputStream is)
    {
        if (is != null)
        {
            try
            {
                is.close();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭输出流
     */
    public void close(OutputStream os)
    {
        if (os != null)
        {
            try
            {
                os.close();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
    }
}