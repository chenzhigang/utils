package com.example.demo.docxtopdf;

import com.example.demo.UncheckBizException;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import com.itextpdf.text.pdf.security.*;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import fr.opensagres.poi.xwpf.converter.core.BasicURIResolver;
import fr.opensagres.poi.xwpf.converter.core.FileImageExtractor;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import fr.opensagres.xdocreport.core.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Entities;
import org.jsoup.select.Elements;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.net.URL;
import java.net.URLConnection;
import java.nio.charset.Charset;
import java.security.KeyStore;
import java.security.PrivateKey;
import java.security.cert.Certificate;

/**
 * 使用 poi + itextpdf 进行word转pdf
 * 先将word转成html，再将html转成pdf
 */
@Slf4j
public class OfficeUtil {

    /**
     * 签名证书地址
     *
     * 使用keytool -genkey命令生成证书（java jdk bin目录下生成）：
     * keytool -genkey -alias signalias -keyalg RSA -keysize 2048 -validity 36500 -keystore sign.keystore
     * testalias是证书别名，可修改为自己想设置的字符，建议使用英文字母和数字
     * test.keystore是证书文件名称，可修改为自己想设置的文件名称，也可以指定完整文件路径
     * 36500是证书的有效期，表示100年有效期，单位天，建议时间设置长一点，避免证书过期
     */
    private static final String KEYSTORE_PATH = "cert/my.keystore";

    /**
     * 证书密码
     */
    private static final String KEYSTORE_PASSWORD = "123456";

    /**
     * 将doc格式文件转成htmlmy.keystore
     *
     * @param docPath  文件路径
     * @param imageDir doc文件中图片存储目录
     * @return html字符串
     */
    public static String doc2Html(String docPath, String imageDir) {
        try {
            return doc2Html(new FileInputStream(docPath), imageDir);
        } catch (FileNotFoundException e) {
            throw new UncheckBizException("doc转html失败");
        }
    }

    /**
     * doc转html
     *
     * @param inputStream 输入流
     * @param imageDir 图片存放目录
     * @return html字符串
     */
    public static String doc2Html(InputStream inputStream, String imageDir) {
        String content = null;
        ByteArrayOutputStream baos = null;
        try {
            HWPFDocument wordDocument = new HWPFDocument(inputStream);
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
            if (StringUtils.isNotEmpty(imageDir)) {
                wordToHtmlConverter.setPicturesManager((content1, pictureType, suggestedName, widthInches, heightInches) -> {
                    File file = new File(imageDir + suggestedName);
                    FileOutputStream fos = null;
                    try {
                        fos = new FileOutputStream(file);
                        fos.write(content1);
                    } catch (IOException e) {
                        log.error("doc转pdf文件输出失败", e);
                    } finally {
                        try {
                            if (fos != null) {
                                fos.close();
                            }
                        } catch (Exception e) {
                            log.error("文件流关闭失败", e);
                        }
                    }
                    return imageDir + suggestedName;
                });
            }
            wordToHtmlConverter.processDocument(wordDocument);
            Document htmlDocument = wordToHtmlConverter.getDocument();
            DOMSource domSource = new DOMSource(htmlDocument);
            baos = new ByteArrayOutputStream();
            StreamResult streamResult = new StreamResult(baos);
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
        } catch (Exception e) {
            log.error("doc转html失败", e);
        } finally {
            try {
                if (baos != null) {
                    content = new String(baos.toByteArray(), "utf-8");
                    baos.close();
                }
                if (inputStream != null) {
                    inputStream.close();
                }
            } catch (Exception e) {
                log.error("关闭文件流失败", e);
            }
        }
        return content;
    }

    /**
     * docx转html
     *
     * @param in 文件输入流
     * @param imageDir  图片所在目录，为空表示不存在
     * @return html字符串
     */
    public static String docx2Html(InputStream in, String imageDir) {
        String content = null;
        ByteArrayOutputStream baos = null;
        try {
            // 1> 加载文档到XWPFDocument
            XWPFDocument document = new XWPFDocument(in);
            // 2> 解析XHTML配置（这里设置IURIResolver来设置图片存放的目录）
            XHTMLOptions options = XHTMLOptions.create(); // 存放word中图片的目录
            if (StringUtils.isNotEmpty(imageDir)) {
                options.setExtractor(new FileImageExtractor(new File(imageDir)));
                options.URIResolver(new BasicURIResolver(imageDir));
            }
            options.setIgnoreStylesIfUnused(false);
            options.setFragment(true);
            // 3> 将XWPFDocument转换成XHTML
            baos = new ByteArrayOutputStream();
            XHTMLConverter.getInstance().convert(document, baos, options);
        } catch (Exception e) {
            log.error("docx转html失败", e);
        } finally {
            try {
                if (in != null) {
                    in.close();
                }
                if (baos != null) {
                    content = new String(baos.toByteArray(), "utf-8");
                    baos.close();
                }
            } catch (Exception e) {
                log.error("文件流关闭失败", e);
            }
        }
        return content;
    }

    /**
     * 将docx格式文件转成html** @param docxPath docx文件路径
     *
     * @param imageDir docx文件中图片存储目录
     * @return html
     */
    public static String docx2Html(String docxPath, String imageDir) {
        try {
            return docx2Html(new FileInputStream(new File(docxPath)), imageDir);
        } catch (FileNotFoundException e) {
            throw new UncheckBizException("docx转html失败");
        }
    }

    /**
     * 使用jsoup规范化html
     *
     * @param html html内容
     * @return 规范化后的html
     */
    private static String formatHtml(String html) {
        org.jsoup.nodes.Document doc = Jsoup.parse(html);
        // 去除过大的宽度
        String style = doc.attr("style");
        if (StringUtils.isNotEmpty(style) && style.contains("width")) {
            doc.attr("style", "");
        }
        Elements divs = doc.select("div");
        for (Element div : divs) {
            String divStyle = div.attr("style");
            if (StringUtils.isNotEmpty(divStyle) && divStyle.contains("width")) {
                div.attr("style", "");
            }
        }
        // jsoup生成闭合标签
        doc.outputSettings().syntax(org.jsoup.nodes.Document.OutputSettings.Syntax.xml);
        doc.outputSettings().escapeMode(Entities.EscapeMode.xhtml);
        return doc.html();
    }

    /**
     * html转成pdf
     *
     * @param html          html字符串
     * @param outputPdfPath 输出pdf路径
     */
    public static void htmlToPdf(String html, String outputPdfPath) {
        com.itextpdf.text.Document document = null;
        ByteArrayInputStream bais = null;
        try {
            document = new com.itextpdf.text.Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPdfPath));
            document.open();
            // html转pdf
            bais = new ByteArrayInputStream(html.getBytes());
            XMLWorkerHelper.getInstance().parseXHtml(writer, document, bais, Charset.forName("UTF-8"), new FontProvider() {
                @Override
                public boolean isRegistered(String s) {
                    return false;
                }
                @Override
                public Font getFont(String s, String s1, boolean embedded, float size, int style, BaseColor baseColor) {
                    // 配置字体
                    Font font = null;
                    try {
                        // 使用jar包：iTextAsian
                        BaseFont bf = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.EMBEDDED);
                        font = new Font(bf, size, style, baseColor);
                        font.setColor(baseColor);
                    } catch (Exception e) {
                        log.error("字体处理异常", e);
                    }
                    return font;
                }
            });
        } catch (Exception e) {
            log.error("html转pdf失败", e);
        } finally {
            if (document != null) {
                document.close();
            }
            if (bais != null) {
                try {
                    bais.close();
                } catch (IOException e) {
                    log.error("文件流关闭失败", e);
                }
            }
        }
    }

    /**
     * word转pdf
     *
     * @param wordPath word地址
     * @param imgDir 图片所在目录
     * @return html字符串
     */
    public static String wordToHtml(String wordPath, String imgDir) {
        if (wordPath.endsWith("docx")) {
            return docx2Html(wordPath, imgDir);
        } else {
            return doc2Html(wordPath, imgDir);
        }
    }

    /**
     * doc转html
     *
     * @param fileType 文件类型：doc、docx
     * @param inputStream 输入流
     * @param imgDir 图片存放目录
     * @return html字符串
     */
    public static String wordToHtml(String fileType, InputStream inputStream, String imgDir) {
        if (fileType.endsWith("docx")) {
            return docx2Html(inputStream, imgDir);
        } else {
            return doc2Html(inputStream, imgDir);
        }
    }

    /**
     * 对pdf进行签章
     *
     * @param pdfReader  pdf文件读取
     * @param imgPath  签章图片路径
     * @param signPage  签章所在的页数，如果为空，默认最后一页
     * @param rOffset  距离页面右侧偏移量
     * @param topOffset  距离页面顶部偏移量
     * @param width  图片宽度
     * @param height  图片高度
     * @param reason   签章原因
     * @param location 签章地点
     * @return pdf字节数组
     */
    public static byte[] sign(PdfReader pdfReader, String imgPath, Integer signPage, float rOffset, float topOffset, float width, float height, String reason, String location) {
        InputStream imgIs = null;
        try {
            imgIs = getInputStreamByAbsPath(imgPath);
            return sign(pdfReader, imgIs, signPage, rOffset, topOffset, width, height, reason, location);
        } catch (FileNotFoundException e) {
            throw new UncheckBizException("签章失败");
        } finally {
            if (null != imgIs) {
                try {
                    imgIs.close();
                } catch (IOException e) {
                    log.error("关闭流失败", e);
                }
            }
        }
    }

    public static byte[] sign(PdfReader pdfReader, InputStream imgIs, Integer signPage, float rOffset, float topOffset, float width, float height, String reason, String location) {
        ByteArrayOutputStream baos = null;
        byte[] resBytes = null;
        try {
            // 读取keystore ，获得私钥和证书链
            KeyStore keyStore = KeyStore.getInstance("JKS");
            keyStore.load(getInputStreamByRelPath(KEYSTORE_PATH), KEYSTORE_PASSWORD.toCharArray());
            String alias = keyStore.aliases().nextElement();
            PrivateKey privateKey = (PrivateKey) keyStore.getKey(alias, KEYSTORE_PASSWORD.toCharArray());
            Certificate[] chain = keyStore.getCertificateChain(alias);

            int totalPage = pdfReader.getNumberOfPages();
            log.info("总页数：{}", totalPage);
            // 传入的为空默认取最后一页
            totalPage = null == signPage ? totalPage : signPage;
            Rectangle rectangle = pdfReader.getPageSize(totalPage);
            float urx = rectangle.getRight() - rOffset;
            float ury = rectangle.getTop() - topOffset;
            float llx = urx - (width + rOffset);
            float lly = ury - (height + topOffset);
            log.info("签名位置：【{},{},{},{}】", urx, ury, llx, lly);
            baos = new ByteArrayOutputStream();
            PdfStamper stamper = PdfStamper.createSignature(pdfReader, baos, 'A', null, true);
            // 获取数字签章属性对象，设定数字签章的属性
            PdfSignatureAppearance appearance = stamper.getSignatureAppearance();
            appearance.setReason(reason);
            appearance.setLocation(location);
            appearance.setVisibleSignature(new Rectangle(llx, lly, urx, ury), totalPage, "sign");
            // 获取盖章图片
            Image image = Image.getInstance(streamToByte(imgIs));
            appearance.setSignatureGraphic(image);
            // 设置认证等级
            appearance.setCertificationLevel(PdfSignatureAppearance.NOT_CERTIFIED);
            // 印章的渲染方式，这里选择只显示印章
            appearance.setRenderingMode(PdfSignatureAppearance.RenderingMode.GRAPHIC);
            ExternalDigest digest = new BouncyCastleDigest();
            // 签名算法，参数依次为：证书秘钥、摘要算法名称，例如MD5 | SHA-1 | SHA-2.... 以及 提供者
            ExternalSignature signature = new PrivateKeySignature(privateKey, DigestAlgorithms.SHA1, null);
            // 调用itext签名方法完成pdf签章
            MakeSignature.signDetached(appearance, digest, signature, chain, null, null, null, 0, MakeSignature.CryptoStandard.CMS);
            resBytes = baos.toByteArray();
        } catch (Exception e) {
            throw new UncheckBizException("pdf文件签章异常", e);
        } finally {
            try {
                if (pdfReader != null) {
                    pdfReader.close();
                }
                if (baos != null) {
                    baos.close();
                }
                if (imgIs != null) {
                    imgIs.close();
                }
            } catch (IOException e) {
                log.error("关闭io流异常", e);
            }
        }
        return resBytes;
    }

    /**
     * 对pdf进行签章
     *
     * @param is      pdf文件输入流
     * @param imgPath  签章图片路径
     * @param signPage  签章所在的页数，如果为空，默认最后一页
     * @param rOffset  距离页面右侧偏移量
     * @param topOffset  距离页面顶部偏移量
     * @param width  图片宽度
     * @param height  图片高度
     * @param reason   签章原因
     * @param location 签章地点
     * @return pdf字节数组
     */
    public static byte[] sign(InputStream is, String imgPath, Integer signPage, float rOffset, float topOffset, float width, float height, String reason, String location) {
        try {
            return sign(new PdfReader(is), imgPath, signPage, rOffset, topOffset, width, height, reason, location);
        } catch (IOException e) {
            throw new UncheckBizException("签名失败");
        } finally {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException e) {
                log.error("关闭流失败", e);
            }
        }
    }

    public static byte[] sign(InputStream is, InputStream imgIs, Integer signPage, float rOffset, float topOffset, float width, float height, String reason, String location) {
        try {
            return sign(new PdfReader(is), imgIs, signPage, rOffset, topOffset, width, height, reason, location);
        } catch (IOException e) {
            throw new UncheckBizException("签名失败");
        } finally {
            try {
                if (null != is) {
                    is.close();
                }
                if (null != imgIs) {
                    imgIs.close();
                }
            } catch (Exception e) {
                log.error("关闭流失败", e);
            }
        }
    }

    /**
     * 对pdf进行签章
     *
     * @param filePath  pdf文件路径
     * @param imgPath  签章图片路径
     * @param signPage  签章所在的页数，如果为空，默认最后一页
     * @param rOffset  距离页面右侧偏移量
     * @param topOffset  距离页面顶部偏移量
     * @param width  图片宽度
     * @param height  图片高度
     * @param reason   签章原因
     * @param location 签章地点
     * @return pdf字节数组
     */
    public static byte[] sign(String filePath, String imgPath, Integer signPage, float rOffset, float topOffset, float width, float height, String reason, String location) throws IOException {
        try {
            return sign(new PdfReader(filePath), imgPath, signPage, rOffset, topOffset, width, height, reason, location);
        } catch (IOException e) {
            throw new UncheckBizException("签名失败");
        }
    }

    public static byte[] sign(String filePath, InputStream imgIs, Integer signPage, float rOffset, float topOffset, float width, float height, String reason, String location) {
        try {
            return sign(new PdfReader(filePath), imgIs, signPage, rOffset, topOffset, width, height, reason, location);
        } catch (IOException e) {
            throw new UncheckBizException("签名失败");
        } finally {
            try {
                if (null != imgIs) {
                    imgIs.close();
                }
            } catch (IOException e) {
                log.error("关闭流失败", e);
            }
        }
    }

    /**
     * 相对路径转输入流
     *
     * @param relPath 相对路径（resources目录下）
     * @return 输入流
     */
    public static InputStream getInputStreamByRelPath(String relPath) {
        return OfficeUtil.class.getClassLoader().getResourceAsStream(relPath);
    }

    /**
     * 相对路径转输入流
     *
     * @param filePath 绝对路径
     * @return 输入流
     */
    public static InputStream getInputStreamByAbsPath(String filePath) throws FileNotFoundException {
        return new FileInputStream(filePath);
    }

    /**
     * 根据文件网络路径获取输入流
     *
     * @param urlPath 网络路径地址
     * @return 输入流
     */
    public static InputStream getInputStreamByUrl(String urlPath) {
        try {
            URL url = new URL(urlPath);
            URLConnection conn = url.openConnection();
            return conn.getInputStream();
        } catch (Exception e) {
            throw new UncheckBizException("读取网络文件失败");
        }
    }

    /**
     * 保存文件
     *
     * @param bytes 输入字节数组
     * @param destPath 目标文件地址（可相对路径也可绝对路径）
     */
    public static void write(byte[] bytes, String destPath) {
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(destPath);
            fos.write(bytes);
            fos.flush();
        } catch (IOException e) {
            log.error("写入数据失败", e);
        } finally {
            if (null != fos) {
                try {
                    fos.close();
                } catch (IOException e) {
                    log.error("关闭输出流失败", e);
                }
            }
        }
    }

    /**
     * 字节数组转流
     *
     * @param bytes 字节数组
     * @return 输入流
     */
    public static InputStream byteToStream(byte[] bytes) {
        return new ByteArrayInputStream(bytes);
    }

    /**
     * @param is 输入流
     * @return 将输入流转为字节数组
     */
    public static byte[] streamToByte(InputStream is) {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        byte[] bytes = new byte[1024];
        int i;
        try {
            while ((i = is.read(bytes)) != -1) {
                bos.write(bytes, 0, i);
            }
        } catch (Exception e) {
            log.error("输入流转字节数组失败", e);
        }
        return bos.toByteArray();
    }


    public static void main(String[] args) throws Exception {
        testNet();
//        testFile();
    }

    private static void testNet() throws Exception {
        String basePath = "D:/wordToPdf/pdf/";
        String pdfPath = basePath + "test1.pdf";

        // 测试doc转pdf
        InputStream is = getInputStreamByUrl("https://images.alpha.pinpianyi.cn//signRebate/contract/template/534d27522c8b4591819997e478e61844.docx");
        String wordHtml = wordToHtml("docx", is, basePath);
        wordHtml = formatHtml(wordHtml);
        System.out.println("pre docHtml->" + wordHtml);
        wordHtml = wordHtml.replace("${1}", "小王八");
        wordHtml = wordHtml.replace("${2}", "录得");
        wordHtml = wordHtml.replace("${3}", "新的");
        wordHtml = wordHtml.replace("${4}", "<img display=\"none\" src=\"D:\\wordToPdf\\pdf\\sign.png\" width=50 height=50/>");
        wordHtml = wordHtml.replace("${sign}", "<img src=\"D:\\wordToPdf\\pdf\\sign.png\" width=50 height=50/>");

        System.out.println("suffix docHtml->" + wordHtml);
        htmlToPdf(wordHtml, pdfPath);

        // pdf加签
        InputStream singIs = getInputStreamByAbsPath(basePath + "/sign.png");
        write(sign(pdfPath, singIs, null, 150, 50, 50, 30, null, null), "sign2.pdf");
    }

    private static void testFile() throws Exception {
        String basePath = "D:/wordToPdf/pdf/";
        String wordPath = basePath + "1.doc";
        String pdfPath = basePath + "test1.pdf";

        // 测试doc转pdf
        String wordHtml = wordToHtml(wordPath, basePath);
        wordHtml = formatHtml(wordHtml);
        System.out.println("pre docHtml->" + wordHtml);
        wordHtml = wordHtml.replace("${1}", "小王八");
        wordHtml = wordHtml.replace("${2}", "录得");
        wordHtml = wordHtml.replace("${3}", "新的");
        wordHtml = wordHtml.replace("${4}", "<img display=\"none\" src=\"D:\\wordToPdf\\pdf\\sign.png\" width=50 height=50/>");
        wordHtml = wordHtml.replace("${sign}", "<img src=\"D:\\wordToPdf\\pdf\\sign.png\" width=50 height=50/>");

        System.out.println("suffix docHtml->" + wordHtml);
        htmlToPdf(wordHtml, pdfPath);
    }

}
