import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;
import org.apache.xmlbeans.XmlObject;
import util.PlanUtil;

import java.awt.geom.Rectangle2D;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;


//CTPicture ct = (CTPicture)shape.getXmlObject();
//                    CTBlipFillProperties bfp = ct.getBlipFill();
//                    bfp.getBlip().setEmbed();


//XSLFPictureData data =  readPPT.addPicture(housePicture,PictureData.PictureType.PNG);
//        new POIXMLDocumentPart.RelationPart()
//        System.out.println(data);

/**
 * created by FingerZhu on 2019/3/7.
 */

public class Plan {

    private XMLSlideShow readPPT;
    private byte[] housePicture;
    private double compassRotation;
    private String path;

    public Plan() throws Exception {
    }

    public void execute() throws Exception {
        path = this.getClass().getClassLoader().getResource("").getPath();

        readPPT = new XMLSlideShow(new FileInputStream(path + "ppt/test.pptx"));
        housePicture = IOUtils.toByteArray(new FileInputStream(path + "picture/house.png"));
        compassRotation = 60;

        List<XSLFSlide> slides = readPPT.getSlides();
        handleOne(slides.get(0));
        handleTwo(slides.get(1));
        handleThree(slides.get(2));
        handleFour(slides.get(3));

        FileOutputStream out = new FileOutputStream("E:\\code\\Workspace\\Java\\ppt\\src\\main\\resources\\ppt\\target.pptx");
        readPPT.write(out);
        readPPT.close();
        out.close();
    }

    private void handleFour(XSLFSlide slide) throws Exception {
        Rectangle2D houseAnchor = null;
        List<XSLFShape> shapes = slide.getShapes();

        // 先处理house
        for (int i = 0; i < shapes.size(); i++) {
            XSLFShape shape = shapes.get(i);
            if (shape.getXmlObject().xmlText().contains("{{house}}")) {
                houseAnchor = shape.getAnchor();
                ((XSLFPictureShape) shape).getPictureData().setData(housePicture);
                break;
            }
        }

        int detailCount = 14;
        int currentDetail = 0;
        List<Integer> deleteShape = new ArrayList<Integer>();
        for (int i = 0; i < shapes.size(); i++) {
            XSLFShape shape = shapes.get(i);
            String text = shape.getXmlObject().xmlText();
            if (shape instanceof XSLFPictureShape) {
                XSLFPictureShape sp = (XSLFPictureShape) shape;
                if (text.contains("{{compass}}")) {
                    sp.setRotation(compassRotation);
                } else if (text.contains("{{detail}}")) {
                    currentDetail += 1;
                    if (currentDetail > detailCount) {
//                        deleteShape.add(i);
                    } else {
                        sp.getPictureData().setData(IOUtils.toByteArray(new FileInputStream(path + "icon/style.png")));
                        // todo
//                        sp.setAnchor(houseAnchor);
                    }
                }
            }
        }

//        for (int i = 0; i < 1; i++) {
//            XSLFShape sp = shapes.get(shapes.size() - 1);
//            XmlObject aa = sp.getXmlObject();
//            System.out.println(aa);
////            slide.removeShape(shapes.get(shapes.size() - 1));
//        }

//        for(int i=0;i<1;i++){
//            slide.removeShape(shapes.get(shapes.size()-1));
//        }

        // 删除一定要倒着删除 否则会报错 应该关系错了
//        for (int i = deleteShape.size() - 1; i >= 0; i--) {
//            slide.removeShape(shapes.get(deleteShape.get(i)));
//        }
    }

    private void handleThree(XSLFSlide slide) {
        List<XSLFShape> shapes = slide.getShapes();
        for (int i = 0; i < shapes.size(); i++) {
            XSLFShape shape = shapes.get(i);
            String text = shape.getXmlObject().xmlText();
            if (shape instanceof XSLFTextShape) {
                XSLFTextShape sp = (XSLFTextShape) shape;
                if (text.contains("{{description}}")) {
                    // todo
                }
            } else if (shape instanceof XSLFPictureShape) {
                XSLFPictureShape sp = (XSLFPictureShape) shape;
                if (text.contains("{{background}}")) {
                    // todo
                }
            }
        }
    }

    private void handleTwo(XSLFSlide slide) throws Exception {
        List<XSLFShape> shapes = slide.getShapes();
        for (int i = 0; i < shapes.size(); i++) {
            XSLFShape shape = shapes.get(i);
            String text = shape.getXmlObject().xmlText();
            if (shape instanceof XSLFPictureShape) {
                XSLFPictureShape sp = (XSLFPictureShape) shape;
                if (text.contains("{{background}}")) {
                    // todo
                }
            }
        }
    }

    private void handleOne(XSLFSlide slide) throws Exception {
        List<XSLFShape> shapes = slide.getShapes();
        for (int i = 0; i < shapes.size(); i++) {
            XSLFShape shape = shapes.get(i);
            String text = shape.getXmlObject().xmlText();
            if (shape instanceof XSLFTextShape) {
                XSLFTextShape sp = (XSLFTextShape) shape;
                if (text.contains("{{title}}")) {
                    PlanUtil.replaceTextShape(sp, "f1F水电费阿萨德佛期望i阿胶怕谁");
                } else if (text.contains("{{author}}")) {
                    PlanUtil.replaceTextShape(sp, "设计师：Finger");
                }
            } else if (shape instanceof XSLFPictureShape) {
                XSLFPictureShape sp = (XSLFPictureShape) shape;
                if (text.contains("{{background}}")) {
                    // todo
//                    sp.getPictureData().setData(housePicture);
                }
            }
        }
    }


}
