package com.cn;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.UUID;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class testPPT2 {
	
	public static void main(String[] args) {

		List<String> list =  converPPTtoImage("C:\\Users\\ThinkPad\\Desktop\\Test.ppt", "C:\\Users\\ThinkPad\\Desktop\\img", "jpg", 20);
		System.out.println(list.get(0));
	}
	
	public static List<String> converPPTtoImage(String orignalPPTFileName, String targetImageFileDir,
	        String imageFormatNameString, int times) {
	    List<String> imgList = new ArrayList<>();
	    List<String> imgNamesList = new ArrayList<String>();// PPTת��ͼƬ���������Ƽ���
	    FileInputStream orignalPPTFileInputStream = null;
	    FileOutputStream orignalPPTFileOutStream = null;
	    XMLSlideShow oneHSLFSlideShow = null;
	    //�����ļ���
	    createDirIfNotExist(targetImageFileDir);
	    try {
	        try {
	            orignalPPTFileInputStream = new FileInputStream(orignalPPTFileName);
	        } catch (FileNotFoundException e) {
	            e.printStackTrace();
	            return Collections.emptyList();
	        }
	        try {
	            oneHSLFSlideShow = new XMLSlideShow(orignalPPTFileInputStream);
	        } catch (IOException e) {
	            e.printStackTrace();
	            return Collections.emptyList();
	        }
	        // ��ȡPPTÿҳ�Ĵ�С����͸߶ȣ�
	        Dimension onePPTPageSize = oneHSLFSlideShow.getPageSize();
	        // ���PPT�ļ��е����е�PPTҳ�棨���ÿһ�Żõ�Ƭ��,��תΪһ���ŵĲ���Ƭ
	        XSLFSlide[] pptPageSlideList = oneHSLFSlideShow.getSlides();
	        // ����ѭ������Ҫ������ʵ�ֶ�PPT�ļ��е�ÿһ�Żõ�Ƭ����ת���Ͳ���
	        for (int i = 0; i < pptPageSlideList.length; i++) {
				setFont(pptPageSlideList[i]);
	            // ����BufferedImage����ͼ��ĳߴ�Ϊԭ����ÿҳ�ĳߴ�*����times
	            BufferedImage oneBufferedImage = new BufferedImage(onePPTPageSize.width * times,
	                    onePPTPageSize.height * times, BufferedImage.TYPE_INT_RGB);
	            Graphics2D oneGraphics2D = oneBufferedImage.createGraphics();
	            // ����ת�����ͼƬ����ɫΪ��ɫ
	            oneGraphics2D.setPaint(Color.white);
	            oneGraphics2D.scale(times, times);// ��ͼƬ�Ŵ�times��
	            oneGraphics2D
	                    .fill(new Rectangle2D.Float(0, 0, onePPTPageSize.width * times, onePPTPageSize.height * times));
	            pptPageSlideList[i].draw(oneGraphics2D);
	            // ����ͼƬ�Ĵ��·����ͼƬ��ʽ��ע�����ɵ�ͼƬ·��Ϊ����·�������ջ�ø���ͼ���ļ�����Ӧ�����������
	            try {
	                String imgName = (i + 1) + "_" + UUID.randomUUID().toString() + "." + imageFormatNameString;
	                imgNamesList.add(imgName);// ��ͼƬ������ӵļ�����
	                imgList.add(imgName);
	                orignalPPTFileOutStream = new FileOutputStream(targetImageFileDir + imgName);

	            } catch (FileNotFoundException e) {
	                e.printStackTrace();
	                return Collections.emptyList();
	            }
	            // ת�����ͼƬ�ļ������ָ����Ŀ¼��
	            try {
	                ImageIO.write(oneBufferedImage, imageFormatNameString, orignalPPTFileOutStream);
	            } catch (IOException e) {
	                e.printStackTrace();
	                return Collections.emptyList();
	            }
	           break;
	        }

	    } finally {
	        try {
	            if (orignalPPTFileInputStream != null) {
	                orignalPPTFileInputStream.close();
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }

	        try {
	            if (orignalPPTFileOutStream != null) {
	                orignalPPTFileOutStream.close();
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }

	    return imgList;
	}
	
	public static File createDirIfNotExist(String file) {
	    File fileDir = new File(file);
	    if (!fileDir.exists()) {
	        fileDir.mkdirs();
	    }
	    return fileDir;
	}
	
	
	private static void setFont(XSLFSlide slide) {
        for (XSLFShape shape : slide.getShapes()){
            if (shape instanceof XSLFTextShape) {
            	XSLFTextShape txtshape = (XSLFTextShape)shape ;
                for (XSLFTextParagraph paragraph : txtshape.getTextParagraphs()) {
                    List<XSLFTextRun> truns = paragraph.getTextRuns();
                    for (XSLFTextRun trun : truns) {
                            trun.setFontFamily("����");
                    }
                }
            }
        }
    }
}