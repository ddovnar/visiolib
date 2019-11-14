package org.dovnard.visio;

import org.w3c.dom.Document;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.IOException;
import java.io.StringReader;

/**
 * User: DovnarDmitriy
 * Date: 08.11.19
 * Time: 14:15
 */
public class DocumentProps implements XmlExportable {
    private String getXmlContent() {
        return "<VisioDocument xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xml:space='preserve'>\n" +
                "                        <DocumentSettings TopPage='0' DefaultTextStyle='3' DefaultLineStyle='3' DefaultFillStyle='3' DefaultGuideStyle='4'>\n" +
                "                            <GlueSettings>9</GlueSettings>\n" +
                "                            <SnapSettings>65847</SnapSettings>\n" +
                "                            <SnapExtensions>34</SnapExtensions>\n" +
                "                            <SnapAngles/>\n" +
                "                            <DynamicGridEnabled>1</DynamicGridEnabled>\n" +
                "                            <ProtectStyles>0</ProtectStyles>\n" +
                "                            <ProtectShapes>0</ProtectShapes>\n" +
                "                            <ProtectMasters>0</ProtectMasters>\n" +
                "                            <ProtectBkgnds>0</ProtectBkgnds>\n" +
                "                        </DocumentSettings>\n" +
                "                        <Colors>\n" +
                "                            <ColorEntry IX='24' RGB='#7F7F7F'/>\n" +
                "                            <ColorEntry IX='25' RGB='#FFFFFF'/>\n" +
                "                        </Colors>\n" +
                "                        <FaceNames>\n" +
                "                            <FaceName NameU='Calibri' UnicodeRanges='-520092929 1073786111 9 0' CharSets='536871327 0' Panose='2 15 5 2 2 2 4 3 2 4' Flags='325'/>\n" +
                "                        </FaceNames>\n" +
                "                        <StyleSheets>\n" +
                "                            <StyleSheet ID='0' NameU='No Style' IsCustomNameU='1' Name='No Style' IsCustomName='1'>\n" +
                "                                <Cell N='EnableLineProps' V='1'/>\n" +
                "                                <Cell N='EnableFillProps' V='1'/>\n" +
                "                                <Cell N='EnableTextProps' V='1'/>\n" +
                "                                <Cell N='HideForApply' V='0'/>\n" +
                "                                <Cell N='LineWeight' V='0.01041666666666667'/>\n" +
                "                                <Cell N='LineColor' V='0'/>\n" +
                "                                <Cell N='LinePattern' V='1'/>\n" +
                "                                <Cell N='Rounding' V='0'/>\n" +
                "                                <Cell N='EndArrowSize' V='2'/>\n" +
                "                                <Cell N='BeginArrow' V='0'/>\n" +
                "                                <Cell N='EndArrow' V='0'/>\n" +
                "                                <Cell N='LineCap' V='0'/>\n" +
                "                                <Cell N='BeginArrowSize' V='2'/>\n" +
                "                                <Cell N='LineColorTrans' V='0'/>\n" +
                "                                <Cell N='CompoundType' V='0'/>\n" +
                "                                <Cell N='FillForegnd' V='1'/>\n" +
                "                                <Cell N='FillBkgnd' V='0'/>\n" +
                "                                <Cell N='FillPattern' V='1'/>\n" +
                "                                <Cell N='ShdwForegnd' V='0'/>\n" +
                "                                <Cell N='ShdwPattern' V='0'/>\n" +
                "                                <Cell N='FillForegndTrans' V='0'/>\n" +
                "                                <Cell N='FillBkgndTrans' V='0'/>\n" +
                "                                <Cell N='ShdwForegndTrans' V='0'/>\n" +
                "                                <Cell N='ShapeShdwType' V='0'/>\n" +
                "                                <Cell N='ShapeShdwOffsetX' V='0'/>\n" +
                "                                <Cell N='ShapeShdwOffsetY' V='0'/>\n" +
                "                                <Cell N='ShapeShdwObliqueAngle' V='0'/>\n" +
                "                                <Cell N='ShapeShdwScaleFactor' V='1'/>\n" +
                "                                <Cell N='ShapeShdwBlur' V='0'/>\n" +
                "                                <Cell N='ShapeShdwShow' V='0'/>\n" +
                "                                <Cell N='LeftMargin' V='0'/>\n" +
                "                                <Cell N='RightMargin' V='0'/>\n" +
                "                                <Cell N='TopMargin' V='0'/>\n" +
                "                                <Cell N='BottomMargin' V='0'/>\n" +
                "                                <Cell N='VerticalAlign' V='1'/>\n" +
                "                                <Cell N='TextBkgnd' V='0'/>\n" +
                "                                <Cell N='DefaultTabStop' V='0.5905511811023622'/>\n" +
                "                                <Cell N='TextDirection' V='0'/>\n" +
                "                                <Cell N='TextBkgndTrans' V='0'/>\n" +
                "                                <Cell N='LockWidth' V='0'/>\n" +
                "                                <Cell N='LockHeight' V='0'/>\n" +
                "                                <Cell N='LockMoveX' V='0'/>\n" +
                "                                <Cell N='LockMoveY' V='0'/>\n" +
                "                                <Cell N='LockAspect' V='0'/>\n" +
                "                                <Cell N='LockDelete' V='0'/>\n" +
                "                                <Cell N='LockBegin' V='0'/>\n" +
                "                                <Cell N='LockEnd' V='0'/>\n" +
                "                                <Cell N='LockRotate' V='0'/>\n" +
                "                                <Cell N='LockCrop' V='0'/>\n" +
                "                                <Cell N='LockVtxEdit' V='0'/>\n" +
                "                                <Cell N='LockTextEdit' V='0'/>\n" +
                "                                <Cell N='LockFormat' V='0'/>\n" +
                "                                <Cell N='LockGroup' V='0'/>\n" +
                "                                <Cell N='LockCalcWH' V='0'/>\n" +
                "                                <Cell N='LockSelect' V='0'/>\n" +
                "                                <Cell N='LockCustProp' V='0'/>\n" +
                "                                <Cell N='LockFromGroupFormat' V='0'/>\n" +
                "                                <Cell N='LockThemeColors' V='0'/>\n" +
                "                                <Cell N='LockThemeEffects' V='0'/>\n" +
                "                                <Cell N='LockThemeConnectors' V='0'/>\n" +
                "                                <Cell N='LockThemeFonts' V='0'/>\n" +
                "                                <Cell N='LockThemeIndex' V='0'/>\n" +
                "                                <Cell N='LockReplace' V='0'/>\n" +
                "                                <Cell N='LockVariation' V='0'/>\n" +
                "                                <Cell N='NoObjHandles' V='0'/>\n" +
                "                                <Cell N='NonPrinting' V='0'/>\n" +
                "                                <Cell N='NoCtlHandles' V='0'/>\n" +
                "                                <Cell N='NoAlignBox' V='0'/>\n" +
                "                                <Cell N='UpdateAlignBox' V='0'/>\n" +
                "                                <Cell N='HideText' V='0'/>\n" +
                "                                <Cell N='DynFeedback' V='0'/>\n" +
                "                                <Cell N='GlueType' V='0'/>\n" +
                "                                <Cell N='WalkPreference' V='0'/>\n" +
                "                                <Cell N='BegTrigger' V='0' F='No Formula'/>\n" +
                "                                <Cell N='EndTrigger' V='0' F='No Formula'/>\n" +
                "                                <Cell N='ObjType' V='0'/>\n" +
                "                                <Cell N='Comment' V=''/>\n" +
                "                                <Cell N='IsDropSource' V='0'/>\n" +
                "                                <Cell N='NoLiveDynamics' V='0'/>\n" +
                "                                <Cell N='LocalizeMerge' V='0'/>\n" +
                "                                <Cell N='NoProofing' V='0'/>\n" +
                "                                <Cell N='Calendar' V='0'/>\n" +
                "                                <Cell N='LangID' V='en-GB'/>\n" +
                "                                <Cell N='ShapeKeywords' V=''/>\n" +
                "                                <Cell N='DropOnPageScale' V='1'/>\n" +
                "                                <Cell N='TheData' V='0' F='No Formula'/>\n" +
                "                                <Cell N='TheText' V='0' F='No Formula'/>\n" +
                "                                <Cell N='EventDblClick' V='0' F='No Formula'/>\n" +
                "                                <Cell N='EventXFMod' V='0' F='No Formula'/>\n" +
                "                                <Cell N='EventDrop' V='0' F='No Formula'/>\n" +
                "                                <Cell N='EventMultiDrop' V='0' F='No Formula'/>\n" +
                "                                <Cell N='HelpTopic' V=''/>\n" +
                "                                <Cell N='Copyright' V=''/>\n" +
                "                                <Cell N='LayerMember' V=''/>\n" +
                "                                <Cell N='XRulerDensity' V='32'/>\n" +
                "                                <Cell N='YRulerDensity' V='32'/>\n" +
                "                                <Cell N='XRulerOrigin' V='0'/>\n" +
                "                                <Cell N='YRulerOrigin' V='0'/>\n" +
                "                                <Cell N='XGridDensity' V='8'/>\n" +
                "                                <Cell N='YGridDensity' V='8'/>\n" +
                "                                <Cell N='XGridSpacing' V='0'/>\n" +
                "                                <Cell N='YGridSpacing' V='0'/>\n" +
                "                                <Cell N='XGridOrigin' V='0'/>\n" +
                "                                <Cell N='YGridOrigin' V='0'/>\n" +
                "                                <Cell N='Gamma' V='1'/>\n" +
                "                                <Cell N='Contrast' V='0.5'/>\n" +
                "                                <Cell N='Brightness' V='0.5'/>\n" +
                "                                <Cell N='Sharpen' V='0'/>\n" +
                "                                <Cell N='Blur' V='0'/>\n" +
                "                                <Cell N='Denoise' V='0'/>\n" +
                "                                <Cell N='Transparency' V='0'/>\n" +
                "                                <Cell N='SelectMode' V='1'/>\n" +
                "                                <Cell N='DisplayMode' V='2'/>\n" +
                "                                <Cell N='IsDropTarget' V='0'/>\n" +
                "                                <Cell N='IsSnapTarget' V='1'/>\n" +
                "                                <Cell N='IsTextEditTarget' V='1'/>\n" +
                "                                <Cell N='DontMoveChildren' V='0'/>\n" +
                "                                <Cell N='ShapePermeableX' V='0'/>\n" +
                "                                <Cell N='ShapePermeableY' V='0'/>\n" +
                "                                <Cell N='ShapePermeablePlace' V='0'/>\n" +
                "                                <Cell N='Relationships' V='0'/>\n" +
                "                                <Cell N='ShapeFixedCode' V='0'/>\n" +
                "                                <Cell N='ShapePlowCode' V='0'/>\n" +
                "                                <Cell N='ShapeRouteStyle' V='0'/>\n" +
                "                                <Cell N='ShapePlaceStyle' V='0'/>\n" +
                "                                <Cell N='ConFixedCode' V='0'/>\n" +
                "                                <Cell N='ConLineJumpCode' V='0'/>\n" +
                "                                <Cell N='ConLineJumpStyle' V='0'/>\n" +
                "                                <Cell N='ConLineJumpDirX' V='0'/>\n" +
                "                                <Cell N='ConLineJumpDirY' V='0'/>\n" +
                "                                <Cell N='ShapePlaceFlip' V='0'/>\n" +
                "                                <Cell N='ConLineRouteExt' V='0'/>\n" +
                "                                <Cell N='ShapeSplit' V='0'/>\n" +
                "                                <Cell N='ShapeSplittable' V='0'/>\n" +
                "                                <Cell N='DisplayLevel' V='0'/>\n" +
                "                                <Cell N='ResizePage' V='0'/>\n" +
                "                                <Cell N='EnableGrid' V='0'/>\n" +
                "                                <Cell N='DynamicsOff' V='0'/>\n" +
                "                                <Cell N='CtrlAsInput' V='0'/>\n" +
                "                                <Cell N='AvoidPageBreaks' V='0'/>\n" +
                "                                <Cell N='PlaceStyle' V='0'/>\n" +
                "                                <Cell N='RouteStyle' V='0'/>\n" +
                "                                <Cell N='PlaceDepth' V='0'/>\n" +
                "                                <Cell N='PlowCode' V='0'/>\n" +
                "                                <Cell N='LineJumpCode' V='1'/>\n" +
                "                                <Cell N='LineJumpStyle' V='0'/>\n" +
                "                                <Cell N='PageLineJumpDirX' V='0'/>\n" +
                "                                <Cell N='PageLineJumpDirY' V='0'/>\n" +
                "                                <Cell N='LineToNodeX' V='0.09842519685039369'/>\n" +
                "                                <Cell N='LineToNodeY' V='0.09842519685039369'/>\n" +
                "                                <Cell N='BlockSizeX' V='0.1968503937007874'/>\n" +
                "                                <Cell N='BlockSizeY' V='0.1968503937007874'/>\n" +
                "                                <Cell N='AvenueSizeX' V='0.2952755905511811'/>\n" +
                "                                <Cell N='AvenueSizeY' V='0.2952755905511811'/>\n" +
                "                                <Cell N='LineToLineX' V='0.09842519685039369'/>\n" +
                "                                <Cell N='LineToLineY' V='0.09842519685039369'/>\n" +
                "                                <Cell N='LineJumpFactorX' V='0.66666666666667'/>\n" +
                "                                <Cell N='LineJumpFactorY' V='0.66666666666667'/>\n" +
                "                                <Cell N='LineAdjustFrom' V='0'/>\n" +
                "                                <Cell N='LineAdjustTo' V='0'/>\n" +
                "                                <Cell N='PlaceFlip' V='0'/>\n" +
                "                                <Cell N='LineRouteExt' V='0'/>\n" +
                "                                <Cell N='PageShapeSplit' V='0'/>\n" +
                "                                <Cell N='PageLeftMargin' V='0.25'/>\n" +
                "                                <Cell N='PageRightMargin' V='0.25'/>\n" +
                "                                <Cell N='PageTopMargin' V='0.25'/>\n" +
                "                                <Cell N='PageBottomMargin' V='0.25'/>\n" +
                "                                <Cell N='ScaleX' V='1'/>\n" +
                "                                <Cell N='ScaleY' V='1'/>\n" +
                "                                <Cell N='PagesX' V='1'/>\n" +
                "                                <Cell N='PagesY' V='1'/>\n" +
                "                                <Cell N='CenterX' V='0'/>\n" +
                "                                <Cell N='CenterY' V='0'/>\n" +
                "                                <Cell N='OnPage' V='0'/>\n" +
                "                                <Cell N='PrintGrid' V='0'/>\n" +
                "                                <Cell N='PrintPageOrientation' V='1'/>\n" +
                "                                <Cell N='PaperKind' V='9'/>\n" +
                "                                <Cell N='PaperSource' V='7'/>\n" +
                "                                <Cell N='QuickStyleLineColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleFillColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleShadowColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleFontColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleLineMatrix' V='100'/>\n" +
                "                                <Cell N='QuickStyleFillMatrix' V='100'/>\n" +
                "                                <Cell N='QuickStyleEffectsMatrix' V='100'/>\n" +
                "                                <Cell N='QuickStyleFontMatrix' V='100'/>\n" +
                "                                <Cell N='QuickStyleType' V='0'/>\n" +
                "                                <Cell N='QuickStyleVariation' V='0'/>\n" +
                "                                <Cell N='LineGradientDir' V='0'/>\n" +
                "                                <Cell N='LineGradientAngle' V='1.5707963267949'/>\n" +
                "                                <Cell N='FillGradientDir' V='0'/>\n" +
                "                                <Cell N='FillGradientAngle' V='1.5707963267949'/>\n" +
                "                                <Cell N='LineGradientEnabled' V='0'/>\n" +
                "                                <Cell N='FillGradientEnabled' V='0'/>\n" +
                "                                <Cell N='RotateGradientWithShape' V='1'/>\n" +
                "                                <Cell N='UseGroupGradient' V='0'/>\n" +
                "                                <Cell N='BevelTopType' V='0'/>\n" +
                "                                <Cell N='BevelTopWidth' V='0'/>\n" +
                "                                <Cell N='BevelTopHeight' V='0'/>\n" +
                "                                <Cell N='BevelBottomType' V='0'/>\n" +
                "                                <Cell N='BevelBottomWidth' V='0'/>\n" +
                "                                <Cell N='BevelBottomHeight' V='0'/>\n" +
                "                                <Cell N='BevelDepthColor' V='1'/>\n" +
                "                                <Cell N='BevelDepthSize' V='0'/>\n" +
                "                                <Cell N='BevelContourColor' V='0'/>\n" +
                "                                <Cell N='BevelContourSize' V='0'/>\n" +
                "                                <Cell N='BevelMaterialType' V='0'/>\n" +
                "                                <Cell N='BevelLightingType' V='0'/>\n" +
                "                                <Cell N='BevelLightingAngle' V='0'/>\n" +
                "                                <Cell N='RotationXAngle' V='0'/>\n" +
                "                                <Cell N='RotationYAngle' V='0'/>\n" +
                "                                <Cell N='RotationZAngle' V='0'/>\n" +
                "                                <Cell N='RotationType' V='0'/>\n" +
                "                                <Cell N='Perspective' V='0'/>\n" +
                "                                <Cell N='DistanceFromGround' V='0'/>\n" +
                "                                <Cell N='KeepTextFlat' V='0'/>\n" +
                "                                <Cell N='ReflectionTrans' V='0'/>\n" +
                "                                <Cell N='ReflectionSize' V='0'/>\n" +
                "                                <Cell N='ReflectionDist' V='0'/>\n" +
                "                                <Cell N='ReflectionBlur' V='0'/>\n" +
                "                                <Cell N='GlowColor' V='1'/>\n" +
                "                                <Cell N='GlowColorTrans' V='0'/>\n" +
                "                                <Cell N='GlowSize' V='0'/>\n" +
                "                                <Cell N='SoftEdgesSize' V='0'/>\n" +
                "                                <Cell N='SketchSeed' V='0'/>\n" +
                "                                <Cell N='SketchEnabled' V='0'/>\n" +
                "                                <Cell N='SketchAmount' V='5'/>\n" +
                "                                <Cell N='SketchLineWeight' V='0.04166666666666666' U='PT'/>\n" +
                "                                <Cell N='SketchLineChange' V='0.14'/>\n" +
                "                                <Cell N='SketchFillChange' V='0.1'/>\n" +
                "                                <Cell N='ColorSchemeIndex' V='0'/>\n" +
                "                                <Cell N='EffectSchemeIndex' V='0'/>\n" +
                "                                <Cell N='ConnectorSchemeIndex' V='0'/>\n" +
                "                                <Cell N='FontSchemeIndex' V='0'/>\n" +
                "                                <Cell N='ThemeIndex' V='0'/>\n" +
                "                                <Cell N='VariationColorIndex' V='0'/>\n" +
                "                                <Cell N='VariationStyleIndex' V='0'/>\n" +
                "                                <Cell N='EmbellishmentIndex' V='0'/>\n" +
                "                                <Cell N='ReplaceLockShapeData' V='0'/>\n" +
                "                                <Cell N='ReplaceLockText' V='0'/>\n" +
                "                                <Cell N='ReplaceLockFormat' V='0'/>\n" +
                "                                <Cell N='ReplaceCopyCells' V='0' U='BOOL' F='No Formula'/>\n" +
                "                                <Cell N='PageWidth' V='0' F='No Formula'/>\n" +
                "                                <Cell N='PageHeight' V='0' F='No Formula'/>\n" +
                "                                <Cell N='ShdwOffsetX' V='0' F='No Formula'/>\n" +
                "                                <Cell N='ShdwOffsetY' V='0' F='No Formula'/>\n" +
                "                                <Cell N='PageScale' V='0' U='MM' F='No Formula'/>\n" +
                "                                <Cell N='DrawingScale' V='0' U='MM' F='No Formula'/>\n" +
                "                                <Cell N='DrawingSizeType' V='0' F='No Formula'/>\n" +
                "                                <Cell N='DrawingScaleType' V='0' F='No Formula'/>\n" +
                "                                <Cell N='InhibitSnap' V='0' F='No Formula'/>\n" +
                "                                <Cell N='PageLockReplace' V='0' U='BOOL' F='No Formula'/>\n" +
                "                                <Cell N='PageLockDuplicate' V='0' U='BOOL' F='No Formula'/>\n" +
                "                                <Cell N='UIVisibility' V='0' F='No Formula'/>\n" +
                "                                <Cell N='ShdwType' V='0' F='No Formula'/>\n" +
                "                                <Cell N='ShdwObliqueAngle' V='0' F='No Formula'/>\n" +
                "                                <Cell N='ShdwScaleFactor' V='0' F='No Formula'/>\n" +
                "                                <Cell N='DrawingResizeType' V='0' F='No Formula'/>\n" +
                "                                <Section N='Character'>\n" +
                "                                    <Row IX='0'>\n" +
                "                                        <Cell N='Font' V='Calibri'/>\n" +
                "                                        <Cell N='Color' V='0'/>\n" +
                "                                        <Cell N='Style' V='0'/>\n" +
                "                                        <Cell N='Case' V='0'/>\n" +
                "                                        <Cell N='Pos' V='0'/>\n" +
                "                                        <Cell N='FontScale' V='1'/>\n" +
                "                                        <Cell N='Size' V='0.1666666666666667'/>\n" +
                "                                        <Cell N='DblUnderline' V='0'/>\n" +
                "                                        <Cell N='Overline' V='0'/>\n" +
                "                                        <Cell N='Strikethru' V='0'/>\n" +
                "                                        <Cell N='DoubleStrikethrough' V='0'/>\n" +
                "                                        <Cell N='Letterspace' V='0'/>\n" +
                "                                        <Cell N='ColorTrans' V='0'/>\n" +
                "                                        <Cell N='AsianFont' V='0'/>\n" +
                "                                        <Cell N='ComplexScriptFont' V='0'/>\n" +
                "                                        <Cell N='ComplexScriptSize' V='-1'/>\n" +
                "                                        <Cell N='LangID' V='en-GB'/>\n" +
                "                                    </Row>\n" +
                "                                </Section>\n" +
                "                                <Section N='Paragraph'>\n" +
                "                                    <Row IX='0'>\n" +
                "                                        <Cell N='IndFirst' V='0'/>\n" +
                "                                        <Cell N='IndLeft' V='0'/>\n" +
                "                                        <Cell N='IndRight' V='0'/>\n" +
                "                                        <Cell N='SpLine' V='-1.2'/>\n" +
                "                                        <Cell N='SpBefore' V='0'/>\n" +
                "                                        <Cell N='SpAfter' V='0'/>\n" +
                "                                        <Cell N='HorzAlign' V='1'/>\n" +
                "                                        <Cell N='Bullet' V='0'/>\n" +
                "                                        <Cell N='BulletStr' V=''/>\n" +
                "                                        <Cell N='BulletFont' V='0'/>\n" +
                "                                        <Cell N='BulletFontSize' V='-1'/>\n" +
                "                                        <Cell N='TextPosAfterBullet' V='0'/>\n" +
                "                                        <Cell N='Flags' V='0'/>\n" +
                "                                    </Row>\n" +
                "                                </Section>\n" +
                "                                <Section N='Tabs'>\n" +
                "                                    <Row IX='0'/>\n" +
                "                                </Section>\n" +
                "                                <Section N='LineGradient'>\n" +
                "                                    <Row IX='0'>\n" +
                "                                        <Cell N='GradientStopColor' V='1'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='0'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='0'/>\n" +
                "                                    </Row>\n" +
                "                                </Section>\n" +
                "                                <Section N='FillGradient'>\n" +
                "                                    <Row IX='0'>\n" +
                "                                        <Cell N='GradientStopColor' V='1'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='0'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='0'/>\n" +
                "                                    </Row>\n" +
                "                                </Section>\n" +
                "                            </StyleSheet>\n" +
                "                            <StyleSheet ID='1' NameU='Text Only' IsCustomNameU='1' Name='Text Only' IsCustomName='1' LineStyle='3' FillStyle='3' TextStyle='3'>\n" +
                "                                <Cell N='EnableLineProps' V='1'/>\n" +
                "                                <Cell N='EnableFillProps' V='1'/>\n" +
                "                                <Cell N='EnableTextProps' V='1'/>\n" +
                "                                <Cell N='HideForApply' V='0'/>\n" +
                "                                <Cell N='LineWeight' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LineColor' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LinePattern' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='Rounding' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='EndArrowSize' V='2' F='Inh'/>\n" +
                "                                <Cell N='BeginArrow' V='0' F='Inh'/>\n" +
                "                                <Cell N='EndArrow' V='0' F='Inh'/>\n" +
                "                                <Cell N='LineCap' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='BeginArrowSize' V='2' F='Inh'/>\n" +
                "                                <Cell N='LineColorTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='CompoundType' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillForegnd' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillBkgnd' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillPattern' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShdwForegnd' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShdwPattern' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillForegndTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillBkgndTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShdwForegndTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwType' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwOffsetX' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwOffsetY' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwObliqueAngle' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwScaleFactor' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwBlur' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwShow' V='0' F='Inh'/>\n" +
                "                                <Cell N='LeftMargin' V='0'/>\n" +
                "                                <Cell N='RightMargin' V='0'/>\n" +
                "                                <Cell N='TopMargin' V='0'/>\n" +
                "                                <Cell N='BottomMargin' V='0'/>\n" +
                "                                <Cell N='VerticalAlign' V='0'/>\n" +
                "                                <Cell N='TextBkgnd' V='0'/>\n" +
                "                                <Cell N='DefaultTabStop' V='0.5905511811023622' F='Inh'/>\n" +
                "                                <Cell N='TextDirection' V='0' F='Inh'/>\n" +
                "                                <Cell N='TextBkgndTrans' V='0' F='Inh'/>\n" +
                "                                <Cell N='LineGradientDir' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LineGradientAngle' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillGradientDir' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillGradientAngle' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LineGradientEnabled' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillGradientEnabled' V='0' F='Inh'/>\n" +
                "                                <Cell N='RotateGradientWithShape' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='UseGroupGradient' V='Themed' F='Inh'/>\n" +
                "                                <Section N='Paragraph'>\n" +
                "                                    <Row IX='0'>\n" +
                "                                        <Cell N='IndFirst' V='0' F='Inh'/>\n" +
                "                                        <Cell N='IndLeft' V='0' F='Inh'/>\n" +
                "                                        <Cell N='IndRight' V='0' F='Inh'/>\n" +
                "                                        <Cell N='SpLine' V='-1.2' F='Inh'/>\n" +
                "                                        <Cell N='SpBefore' V='0' F='Inh'/>\n" +
                "                                        <Cell N='SpAfter' V='0' F='Inh'/>\n" +
                "                                        <Cell N='HorzAlign' V='0'/>\n" +
                "                                        <Cell N='Bullet' V='0' F='Inh'/>\n" +
                "                                        <Cell N='BulletStr' V='' F='Inh'/>\n" +
                "                                        <Cell N='BulletFont' V='0' F='Inh'/>\n" +
                "                                        <Cell N='BulletFontSize' V='-1' F='Inh'/>\n" +
                "                                        <Cell N='TextPosAfterBullet' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Flags' V='0' F='Inh'/>\n" +
                "                                    </Row>\n" +
                "                                </Section>\n" +
                "                            </StyleSheet>\n" +
                "                            <StyleSheet ID='2' NameU='None' IsCustomNameU='1' Name='None' IsCustomName='1' LineStyle='3' FillStyle='3' TextStyle='3'>\n" +
                "                                <Cell N='EnableLineProps' V='1'/>\n" +
                "                                <Cell N='EnableFillProps' V='1'/>\n" +
                "                                <Cell N='EnableTextProps' V='1'/>\n" +
                "                                <Cell N='HideForApply' V='0'/>\n" +
                "                                <Cell N='LineWeight' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LineColor' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LinePattern' V='0'/>\n" +
                "                                <Cell N='Rounding' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='EndArrowSize' V='2' F='Inh'/>\n" +
                "                                <Cell N='BeginArrow' V='0' F='Inh'/>\n" +
                "                                <Cell N='EndArrow' V='0' F='Inh'/>\n" +
                "                                <Cell N='LineCap' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='BeginArrowSize' V='2' F='Inh'/>\n" +
                "                                <Cell N='LineColorTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='CompoundType' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillForegnd' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillBkgnd' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillPattern' V='0'/>\n" +
                "                                <Cell N='ShdwForegnd' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShdwPattern' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillForegndTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillBkgndTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShdwForegndTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwType' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwOffsetX' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwOffsetY' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwObliqueAngle' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwScaleFactor' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwBlur' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwShow' V='0' F='Inh'/>\n" +
                "                                <Cell N='LineGradientDir' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LineGradientAngle' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillGradientDir' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillGradientAngle' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LineGradientEnabled' V='0'/>\n" +
                "                                <Cell N='FillGradientEnabled' V='0'/>\n" +
                "                                <Cell N='RotateGradientWithShape' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='UseGroupGradient' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='QuickStyleLineColor' V='100' F='Inh'/>\n" +
                "                                <Cell N='QuickStyleFillColor' V='100' F='Inh'/>\n" +
                "                                <Cell N='QuickStyleShadowColor' V='100' F='Inh'/>\n" +
                "                                <Cell N='QuickStyleFontColor' V='100' F='Inh'/>\n" +
                "                                <Cell N='QuickStyleLineMatrix' V='100' F='Inh'/>\n" +
                "                                <Cell N='QuickStyleFillMatrix' V='100' F='Inh'/>\n" +
                "                                <Cell N='QuickStyleEffectsMatrix' V='0' F='GUARD(0)'/>\n" +
                "                                <Cell N='QuickStyleFontMatrix' V='100' F='Inh'/>\n" +
                "                                <Cell N='QuickStyleType' V='0' F='Inh'/>\n" +
                "                                <Cell N='QuickStyleVariation' V='2'/>\n" +
                "                            </StyleSheet>\n" +
                "                            <StyleSheet ID='3' NameU='Normal' IsCustomNameU='1' Name='Normal' IsCustomName='1' LineStyle='6' FillStyle='6' TextStyle='6'>\n" +
                "                                <Cell N='EnableLineProps' V='1'/>\n" +
                "                                <Cell N='EnableFillProps' V='1'/>\n" +
                "                                <Cell N='EnableTextProps' V='1'/>\n" +
                "                                <Cell N='HideForApply' V='0'/>\n" +
                "                                <Cell N='LeftMargin' V='0.05555555555555555' U='PT'/>\n" +
                "                                <Cell N='RightMargin' V='0.05555555555555555' U='PT'/>\n" +
                "                                <Cell N='TopMargin' V='0.05555555555555555' U='PT'/>\n" +
                "                                <Cell N='BottomMargin' V='0.05555555555555555' U='PT'/>\n" +
                "                                <Cell N='VerticalAlign' V='1' F='Inh'/>\n" +
                "                                <Cell N='TextBkgnd' V='0' F='Inh'/>\n" +
                "                                <Cell N='DefaultTabStop' V='0.5905511811023622' F='Inh'/>\n" +
                "                                <Cell N='TextDirection' V='0' F='Inh'/>\n" +
                "                                <Cell N='TextBkgndTrans' V='0' F='Inh'/>\n" +
                "                            </StyleSheet>\n" +
                "                            <StyleSheet ID='4' NameU='Guide' IsCustomNameU='1' Name='Guide' IsCustomName='1' LineStyle='3' FillStyle='3' TextStyle='3'>\n" +
                "                                <Cell N='EnableLineProps' V='1'/>\n" +
                "                                <Cell N='EnableFillProps' V='1'/>\n" +
                "                                <Cell N='EnableTextProps' V='1'/>\n" +
                "                                <Cell N='HideForApply' V='0'/>\n" +
                "                                <Cell N='LineWeight' V='0' U='PT'/>\n" +
                "                                <Cell N='LineColor' V='#7f7f7f'/>\n" +
                "                                <Cell N='LinePattern' V='23'/>\n" +
                "                                <Cell N='Rounding' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='EndArrowSize' V='2' F='Inh'/>\n" +
                "                                <Cell N='BeginArrow' V='0' F='Inh'/>\n" +
                "                                <Cell N='EndArrow' V='0' F='Inh'/>\n" +
                "                                <Cell N='LineCap' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='BeginArrowSize' V='2' F='Inh'/>\n" +
                "                                <Cell N='LineColorTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='CompoundType' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillForegnd' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillBkgnd' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillPattern' V='0'/>\n" +
                "                                <Cell N='ShdwForegnd' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShdwPattern' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillForegndTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillBkgndTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShdwForegndTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwType' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwOffsetX' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwOffsetY' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwObliqueAngle' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwScaleFactor' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwBlur' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='ShapeShdwShow' V='0' F='Inh'/>\n" +
                "                                <Cell N='LineGradientDir' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LineGradientAngle' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillGradientDir' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='FillGradientAngle' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LineGradientEnabled' V='0'/>\n" +
                "                                <Cell N='FillGradientEnabled' V='0'/>\n" +
                "                                <Cell N='RotateGradientWithShape' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='UseGroupGradient' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LeftMargin' V='0.05555555555555555' U='PT' F='Inh'/>\n" +
                "                                <Cell N='RightMargin' V='0.05555555555555555' U='PT' F='Inh'/>\n" +
                "                                <Cell N='TopMargin' V='0'/>\n" +
                "                                <Cell N='BottomMargin' V='0'/>\n" +
                "                                <Cell N='VerticalAlign' V='2'/>\n" +
                "                                <Cell N='TextBkgnd' V='0' F='Inh'/>\n" +
                "                                <Cell N='DefaultTabStop' V='0.5905511811023622' F='Inh'/>\n" +
                "                                <Cell N='TextDirection' V='0' F='Inh'/>\n" +
                "                                <Cell N='TextBkgndTrans' V='0' F='Inh'/>\n" +
                "                                <Cell N='NoObjHandles' V='0' F='Inh'/>\n" +
                "                                <Cell N='NonPrinting' V='1'/>\n" +
                "                                <Cell N='NoCtlHandles' V='0' F='Inh'/>\n" +
                "                                <Cell N='NoAlignBox' V='0' F='Inh'/>\n" +
                "                                <Cell N='UpdateAlignBox' V='0' F='Inh'/>\n" +
                "                                <Cell N='HideText' V='0' F='Inh'/>\n" +
                "                                <Cell N='DynFeedback' V='0' F='Inh'/>\n" +
                "                                <Cell N='GlueType' V='0' F='Inh'/>\n" +
                "                                <Cell N='WalkPreference' V='0' F='Inh'/>\n" +
                "                                <Cell N='BegTrigger' V='0' F='No Formula'/>\n" +
                "                                <Cell N='EndTrigger' V='0' F='No Formula'/>\n" +
                "                                <Cell N='ObjType' V='0' F='Inh'/>\n" +
                "                                <Cell N='Comment' V='' F='Inh'/>\n" +
                "                                <Cell N='IsDropSource' V='0' F='Inh'/>\n" +
                "                                <Cell N='NoLiveDynamics' V='0' F='Inh'/>\n" +
                "                                <Cell N='LocalizeMerge' V='0' F='Inh'/>\n" +
                "                                <Cell N='NoProofing' V='0' F='Inh'/>\n" +
                "                                <Cell N='Calendar' V='0' F='Inh'/>\n" +
                "                                <Cell N='LangID' V='en-GB' F='Inh'/>\n" +
                "                                <Cell N='ShapeKeywords' V='' F='Inh'/>\n" +
                "                                <Cell N='DropOnPageScale' V='1' F='Inh'/>\n" +
                "                                <Cell N='ShapePermeableX' V='1'/>\n" +
                "                                <Cell N='ShapePermeableY' V='1'/>\n" +
                "                                <Cell N='ShapePermeablePlace' V='1'/>\n" +
                "                                <Cell N='Relationships' V='0' F='Inh'/>\n" +
                "                                <Cell N='ShapeFixedCode' V='0' F='Inh'/>\n" +
                "                                <Cell N='ShapePlowCode' V='0' F='Inh'/>\n" +
                "                                <Cell N='ShapeRouteStyle' V='0' F='Inh'/>\n" +
                "                                <Cell N='ShapePlaceStyle' V='0' F='Inh'/>\n" +
                "                                <Cell N='ConFixedCode' V='0' F='Inh'/>\n" +
                "                                <Cell N='ConLineJumpCode' V='0' F='Inh'/>\n" +
                "                                <Cell N='ConLineJumpStyle' V='0' F='Inh'/>\n" +
                "                                <Cell N='ConLineJumpDirX' V='0' F='Inh'/>\n" +
                "                                <Cell N='ConLineJumpDirY' V='0' F='Inh'/>\n" +
                "                                <Cell N='ShapePlaceFlip' V='0' F='Inh'/>\n" +
                "                                <Cell N='ConLineRouteExt' V='0' F='Inh'/>\n" +
                "                                <Cell N='ShapeSplit' V='0' F='Inh'/>\n" +
                "                                <Cell N='ShapeSplittable' V='0' F='Inh'/>\n" +
                "                                <Cell N='DisplayLevel' V='0' F='Inh'/>\n" +
                "                                <Section N='Character'>\n" +
                "                                    <Row IX='0'>\n" +
                "                                        <Cell N='Font' V='Themed' F='Inh'/>\n" +
                "                                        <Cell N='Color' V='4'/>\n" +
                "                                        <Cell N='Style' V='Themed' F='Inh'/>\n" +
                "                                        <Cell N='Case' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Pos' V='0' F='Inh'/>\n" +
                "                                        <Cell N='FontScale' V='1' F='Inh'/>\n" +
                "                                        <Cell N='Size' V='0.125'/>\n" +
                "                                        <Cell N='DblUnderline' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Overline' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Strikethru' V='0' F='Inh'/>\n" +
                "                                        <Cell N='DoubleStrikethrough' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Letterspace' V='0' F='Inh'/>\n" +
                "                                        <Cell N='ColorTrans' V='0' F='Inh'/>\n" +
                "                                        <Cell N='AsianFont' V='Themed' F='Inh'/>\n" +
                "                                        <Cell N='ComplexScriptFont' V='Themed' F='Inh'/>\n" +
                "                                        <Cell N='ComplexScriptSize' V='-1' F='Inh'/>\n" +
                "                                        <Cell N='LangID' V='en-GB' F='Inh'/>\n" +
                "                                    </Row>\n" +
                "                                </Section>\n" +
                "                            </StyleSheet>\n" +
                "                            <StyleSheet ID='6' NameU='Theme' IsCustomNameU='1' Name='Theme' IsCustomName='1' LineStyle='0' FillStyle='0' TextStyle='0'>\n" +
                "                                <Cell N='EnableLineProps' V='1'/>\n" +
                "                                <Cell N='EnableFillProps' V='1'/>\n" +
                "                                <Cell N='EnableTextProps' V='1'/>\n" +
                "                                <Cell N='HideForApply' V='0'/>\n" +
                "                                <Cell N='LineWeight' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='LineColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='LinePattern' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='Rounding' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='EndArrowSize' V='2' F='Inh'/>\n" +
                "                                <Cell N='BeginArrow' V='0' F='Inh'/>\n" +
                "                                <Cell N='EndArrow' V='0' F='Inh'/>\n" +
                "                                <Cell N='LineCap' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='BeginArrowSize' V='2' F='Inh'/>\n" +
                "                                <Cell N='LineColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='CompoundType' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='FillForegnd' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='FillBkgnd' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='FillPattern' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ShdwForegnd' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ShdwPattern' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='FillForegndTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='FillBkgndTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ShdwForegndTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ShapeShdwType' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ShapeShdwOffsetX' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ShapeShdwOffsetY' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ShapeShdwObliqueAngle' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ShapeShdwScaleFactor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ShapeShdwBlur' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ShapeShdwShow' V='0' F='Inh'/>\n" +
                "                                <Cell N='LineGradientDir' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='LineGradientAngle' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='FillGradientDir' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='FillGradientAngle' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='LineGradientEnabled' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='FillGradientEnabled' V='0'/>\n" +
                "                                <Cell N='RotateGradientWithShape' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='UseGroupGradient' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='BevelTopType' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='BevelTopWidth' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='BevelTopHeight' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='BevelBottomType' V='0' F='Inh'/>\n" +
                "                                <Cell N='BevelBottomWidth' V='0' F='Inh'/>\n" +
                "                                <Cell N='BevelBottomHeight' V='0' F='Inh'/>\n" +
                "                                <Cell N='BevelDepthColor' V='1' F='Inh'/>\n" +
                "                                <Cell N='BevelDepthSize' V='0' F='Inh'/>\n" +
                "                                <Cell N='BevelContourColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='BevelContourSize' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='BevelMaterialType' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='BevelLightingType' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='BevelLightingAngle' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ReflectionTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ReflectionSize' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ReflectionDist' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='ReflectionBlur' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='GlowColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='GlowColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='GlowSize' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='SoftEdgesSize' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='SketchSeed' V='0' F='Inh'/>\n" +
                "                                <Cell N='SketchEnabled' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='SketchAmount' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='SketchLineWeight' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='SketchLineChange' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='SketchFillChange' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='QuickStyleLineColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleFillColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleShadowColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleFontColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleLineMatrix' V='100'/>\n" +
                "                                <Cell N='QuickStyleFillMatrix' V='100'/>\n" +
                "                                <Cell N='QuickStyleEffectsMatrix' V='100'/>\n" +
                "                                <Cell N='QuickStyleFontMatrix' V='100'/>\n" +
                "                                <Cell N='QuickStyleType' V='0' F='Inh'/>\n" +
                "                                <Cell N='QuickStyleVariation' V='0' F='Inh'/>\n" +
                "                                <Cell N='ColorSchemeIndex' V='65534'/>\n" +
                "                                <Cell N='EffectSchemeIndex' V='65534'/>\n" +
                "                                <Cell N='ConnectorSchemeIndex' V='65534'/>\n" +
                "                                <Cell N='FontSchemeIndex' V='65534'/>\n" +
                "                                <Cell N='ThemeIndex' V='65534'/>\n" +
                "                                <Cell N='VariationColorIndex' V='65534'/>\n" +
                "                                <Cell N='VariationStyleIndex' V='65534'/>\n" +
                "                                <Cell N='EmbellishmentIndex' V='65534'/>\n" +
                "                                <Section N='Character'>\n" +
                "                                    <Row IX='0'>\n" +
                "                                        <Cell N='Font' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='Color' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='Style' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='Case' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Pos' V='0' F='Inh'/>\n" +
                "                                        <Cell N='FontScale' V='1' F='Inh'/>\n" +
                "                                        <Cell N='Size' V='0.1666666666666667' F='Inh'/>\n" +
                "                                        <Cell N='DblUnderline' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Overline' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Strikethru' V='0' F='Inh'/>\n" +
                "                                        <Cell N='DoubleStrikethrough' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Letterspace' V='0' F='Inh'/>\n" +
                "                                        <Cell N='ColorTrans' V='0' F='Inh'/>\n" +
                "                                        <Cell N='AsianFont' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='ComplexScriptFont' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='ComplexScriptSize' V='-1' F='Inh'/>\n" +
                "                                        <Cell N='LangID' V='en-GB' F='Inh'/>\n" +
                "                                    </Row>\n" +
                "                                </Section>\n" +
                "                                <Section N='FillGradient'>\n" +
                "                                    <Row IX='0'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='1'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='2'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='3'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='4'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='5'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='6'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='7'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='8'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='9'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                </Section>\n" +
                "                                <Section N='LineGradient'>\n" +
                "                                    <Row IX='0'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='1'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='2'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='3'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='4'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='5'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='6'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='7'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='8'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                    <Row IX='9'>\n" +
                "                                        <Cell N='GradientStopColor' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopColorTrans' V='Themed' F='THEMEVAL()'/>\n" +
                "                                        <Cell N='GradientStopPosition' V='Themed' F='THEMEVAL()'/>\n" +
                "                                    </Row>\n" +
                "                                </Section>\n" +
                "                            </StyleSheet>\n" +
                "                            <StyleSheet ID='7' NameU='Connector' IsCustomNameU='1' Name='Connector' IsCustomName='1' LineStyle='3' FillStyle='3' TextStyle='3'>\n" +
                "                                <Cell N='EnableLineProps' V='1'/>\n" +
                "                                <Cell N='EnableFillProps' V='1'/>\n" +
                "                                <Cell N='EnableTextProps' V='1'/>\n" +
                "                                <Cell N='HideForApply' V='0'/>\n" +
                "                                <Cell N='LeftMargin' V='0.05555555555555555' U='PT' F='Inh'/>\n" +
                "                                <Cell N='RightMargin' V='0.05555555555555555' U='PT' F='Inh'/>\n" +
                "                                <Cell N='TopMargin' V='0.05555555555555555' U='PT' F='Inh'/>\n" +
                "                                <Cell N='BottomMargin' V='0.05555555555555555' U='PT' F='Inh'/>\n" +
                "                                <Cell N='VerticalAlign' V='1' F='Inh'/>\n" +
                "                                <Cell N='TextBkgnd' V='#ffffff' F='THEMEGUARD(THEMEVAL(\"BackgroundColor\")+1)'/>\n" +
                "                                <Cell N='DefaultTabStop' V='0.5905511811023622' F='Inh'/>\n" +
                "                                <Cell N='TextDirection' V='0' F='Inh'/>\n" +
                "                                <Cell N='TextBkgndTrans' V='0' F='Inh'/>\n" +
                "                                <Cell N='QuickStyleLineColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleFillColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleShadowColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleFontColor' V='100'/>\n" +
                "                                <Cell N='QuickStyleLineMatrix' V='1'/>\n" +
                "                                <Cell N='QuickStyleFillMatrix' V='1'/>\n" +
                "                                <Cell N='QuickStyleEffectsMatrix' V='1'/>\n" +
                "                                <Cell N='QuickStyleFontMatrix' V='1'/>\n" +
                "                                <Cell N='QuickStyleType' V='0'/>\n" +
                "                                <Cell N='QuickStyleVariation' V='0'/>\n" +
                "                                <Cell N='LineWeight' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LineColor' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='LinePattern' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='Rounding' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='EndArrowSize' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='BeginArrow' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='EndArrow' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='LineCap' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='BeginArrowSize' V='Themed' F='THEMEVAL()'/>\n" +
                "                                <Cell N='LineColorTrans' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='CompoundType' V='Themed' F='Inh'/>\n" +
                "                                <Cell N='NoObjHandles' V='0' F='Inh'/>\n" +
                "                                <Cell N='NonPrinting' V='0' F='Inh'/>\n" +
                "                                <Cell N='NoCtlHandles' V='0' F='Inh'/>\n" +
                "                                <Cell N='NoAlignBox' V='0' F='Inh'/>\n" +
                "                                <Cell N='UpdateAlignBox' V='0' F='Inh'/>\n" +
                "                                <Cell N='HideText' V='0' F='Inh'/>\n" +
                "                                <Cell N='DynFeedback' V='0' F='Inh'/>\n" +
                "                                <Cell N='GlueType' V='0' F='Inh'/>\n" +
                "                                <Cell N='WalkPreference' V='0' F='Inh'/>\n" +
                "                                <Cell N='BegTrigger' V='0' F='No Formula'/>\n" +
                "                                <Cell N='EndTrigger' V='0' F='No Formula'/>\n" +
                "                                <Cell N='ObjType' V='0' F='Inh'/>\n" +
                "                                <Cell N='Comment' V='' F='Inh'/>\n" +
                "                                <Cell N='IsDropSource' V='0' F='Inh'/>\n" +
                "                                <Cell N='NoLiveDynamics' V='0' F='Inh'/>\n" +
                "                                <Cell N='LocalizeMerge' V='0' F='Inh'/>\n" +
                "                                <Cell N='NoProofing' V='0' F='Inh'/>\n" +
                "                                <Cell N='Calendar' V='0' F='Inh'/>\n" +
                "                                <Cell N='LangID' V='en-GB' F='Inh'/>\n" +
                "                                <Cell N='ShapeKeywords' V='' F='Inh'/>\n" +
                "                                <Cell N='DropOnPageScale' V='1' F='Inh'/>\n" +
                "                                <Section N='Character'>\n" +
                "                                    <Row IX='0'>\n" +
                "                                        <Cell N='Font' V='Themed' F='Inh'/>\n" +
                "                                        <Cell N='Color' V='Themed' F='Inh'/>\n" +
                "                                        <Cell N='Style' V='Themed' F='Inh'/>\n" +
                "                                        <Cell N='Case' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Pos' V='0' F='Inh'/>\n" +
                "                                        <Cell N='FontScale' V='1' F='Inh'/>\n" +
                "                                        <Cell N='Size' V='0.1111111111111111'/>\n" +
                "                                        <Cell N='DblUnderline' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Overline' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Strikethru' V='0' F='Inh'/>\n" +
                "                                        <Cell N='DoubleStrikethrough' V='0' F='Inh'/>\n" +
                "                                        <Cell N='Letterspace' V='0' F='Inh'/>\n" +
                "                                        <Cell N='ColorTrans' V='0' F='Inh'/>\n" +
                "                                        <Cell N='AsianFont' V='Themed' F='Inh'/>\n" +
                "                                        <Cell N='ComplexScriptFont' V='Themed' F='Inh'/>\n" +
                "                                        <Cell N='ComplexScriptSize' V='-1' F='Inh'/>\n" +
                "                                        <Cell N='LangID' V='en-GB' F='Inh'/>\n" +
                "                                    </Row>\n" +
                "                                </Section>\n" +
                "                            </StyleSheet>\n" +
                "                        </StyleSheets>\n" +
                "                        <DocumentSheet NameU='TheDoc' IsCustomNameU='1' Name='TheDoc' IsCustomName='1' LineStyle='0' FillStyle='0' TextStyle='0'>\n" +
                "                            <Cell N='OutputFormat' V='0'/>\n" +
                "                            <Cell N='LockPreview' V='0'/>\n" +
                "                            <Cell N='AddMarkup' V='0'/>\n" +
                "                            <Cell N='ViewMarkup' V='0'/>\n" +
                "                            <Cell N='DocLockReplace' V='0' U='BOOL'/>\n" +
                "                            <Cell N='NoCoauth' V='0' U='BOOL'/>\n" +
                "                            <Cell N='DocLockDuplicatePage' V='0' U='BOOL'/>\n" +
                "                            <Cell N='PreviewQuality' V='0'/>\n" +
                "                            <Cell N='PreviewScope' V='0'/>\n" +
                "                            <Cell N='DocLangID' V='en-US'/>\n" +
                "                            <Section N='User'>\n" +
                "                                <Row N='msvNoAutoConnect'>\n" +
                "                                    <Cell N='Value' V='1'/>\n" +
                "                                    <Cell N='Prompt' V='' F='No Formula'/>\n" +
                "                                </Row>\n" +
                "                            </Section>\n" +
                "                        </DocumentSheet>\n" +
                "                    </VisioDocument>";
    }
    @Override
    public Document toXml() throws ParserConfigurationException {
        try {
            DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

            return docBuilder.parse(new InputSource(new StringReader(getXmlContent())));
        } catch (SAXException e) {
            throw new ParserConfigurationException();
        } catch (IOException e) {
            throw new ParserConfigurationException();
        }
    }

    @Override
    public String getFileName() {
        return "/visio/document.xml";
    }

    @Override
    public void writeToFile(String path) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, getFileName(), toXml());
    }

    @Override
    public void writeToFile(String path, String filename) throws IOException, TransformerException, ParserConfigurationException {
        XmlFileWriter.writeToFile(path, filename, toXml());
    }
}
