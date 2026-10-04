import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XContentEnumerationAccess;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.text.*;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import org.w3c.dom.Element;
import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import java.nio.file.*;
import java.util.*;
import java.util.regex.Pattern;
import java.util.zip.ZipFile;

/** Explicit Linux PDF compatibility export. Never saves or rewrites the source DOCX. */
public final class LibreOfficeBaselineExporter {
    private static final String W="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final String V="urn:schemas-microsoft-com:vml";
    private static final String O="urn:schemas-microsoft-com:office:office";
    private record Metric(String shapeId, double widthPt, double heightPt, int positionHalfPt) {}
    private record FontPortion(XPropertySet properties, XTextRange range, String requested, String text) {}
    enum FontAvailability { INSTALLED, MISSING, UNRESOLVED_FAMILY }
    // Deliberately bounded aliases. Unknown family names fail closed rather
    // than being overwritten on the strength of a display-name-only font list.
    // Canonicalization follows LibreOffice's fontdefs.cxx; KaiTi's legacy GB2312
    // family is separate from KaiTi/SimKai, never a guessed equivalent.
    private static final Map<String,String> FONT_ALIASES=Map.ofEntries(
        Map.entry("宋体","simsun"),Map.entry("新宋体","nsimsun"),Map.entry("黑体","simhei"),
        Map.entry("楷体","kaiti"),Map.entry("simkai","kaiti"),Map.entry("楷体gb2312","kaitigb2312"),
        Map.entry("仿宋","fangsong"),Map.entry("仿宋gb2312","fangsonggb2312"),
        Map.entry("細明體","mingliu"),Map.entry("新細明體","pmingliu"),
        Map.entry("微軟正黑體","microsoftjhenghei"),Map.entry("微软雅黑","microsoftyahei"),
        Map.entry("msゴシック","msgothic"),Map.entry("mspゴシック","mspgothic"),
        Map.entry("ms明朝","msmincho"),Map.entry("msp明朝","mspmincho"),Map.entry("メイリオ","meiryo"),
        Map.entry("바탕","batang"),Map.entry("바탕체","batangche"),Map.entry("굴림","gulim"),
        Map.entry("굴림체","gulimche"),Map.entry("돋움","dotum"),Map.entry("돋움체","dotumche"));
    private static final class FontDependencyException extends RuntimeException {
        FontDependencyException(String message) { super(message); }
    }
    private static PropertyValue property(String name,Object value){var p=new PropertyValue();p.Name=name;p.Value=value;return p;}
    private static void require(boolean condition,String message){if(!condition)throw new IllegalStateException(message);}

    public static void main(String[] args) {
        try { run(args); System.exit(0); }
        catch (FontDependencyException failure) { System.err.println(failure.getMessage()); System.exit(69); }
        catch (Exception failure) { failure.printStackTrace(); System.exit(1); }
    }
    private static void run(String[] args)throws Exception {
        require(args.length==5 || args.length==6,"usage: <UNO-port> <source.docx> <native.pdf> <compatible.pdf> <audit.tsv> [fallback-CJK-font]");
        Path source=Path.of(args[1]).toAbsolutePath();
        var metrics=readMetrics(source);
        var local=Bootstrap.createInitialComponentContext(null);
        var resolver=UnoRuntime.queryInterface(XUnoUrlResolver.class,local.getServiceManager().createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local));
        XComponentContext context=null;
        for(int i=0;i<150;i++)try{
            context=UnoRuntime.queryInterface(XComponentContext.class,resolver.resolve("uno:socket,host=127.0.0.1,port="+Integer.parseInt(args[0])+";urp;StarOffice.ComponentContext"));break;
        }catch(Exception unavailable){Thread.sleep(100);}
        require(context!=null,"Could not connect to the private loopback LibreOffice instance");
        var desktop=context.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop",context);
        var loader=UnoRuntime.queryInterface(XComponentLoader.class,desktop);
        XComponent document=loader.loadComponentFromURL(source.toUri().toString(),"_blank",0,new PropertyValue[]{
            property("Hidden",true),property("ReadOnly",true),
            property("MacroExecutionMode",com.sun.star.document.MacroExecMode.NEVER_EXECUTE),
            property("UpdateDocMode",com.sun.star.document.UpdateDocMode.NO_UPDATE)});
        require(document!=null,"LibreOffice did not open the DOCX");
        try {
            var objects=UnoRuntime.queryInterface(XTextEmbeddedObjectsSupplier.class,document).getEmbeddedObjects();
            Set<String> expectedNames=new HashSet<>(Arrays.asList(objects.getElementNames()));
            List<String> ordered=new ArrayList<>();
            visitText(UnoRuntime.queryInterface(XTextDocument.class,document).getText(),expectedNames,ordered);
            require(ordered.size()==metrics.size() && ordered.size()==expectedNames.size(),
                "OLE count/order coverage mismatch: XML="+metrics.size()+", body="+ordered.size()+", UNO="+expectedNames.size());
            require(new HashSet<>(ordered).size()==ordered.size(),"An OLE was encountered more than once");
            List<XPropertySet> properties=new ArrayList<>();
            List<int[]> originalObjectSizes=new ArrayList<>();
            for(int i=0;i<ordered.size();i++){
                var p=UnoRuntime.queryInterface(XPropertySet.class,objects.getByName(ordered.get(i)));var m=metrics.get(i);
                require(TextContentAnchorType.AS_CHARACTER.equals(p.getPropertyValue("AnchorType")),"Only inline OLEs are supported: "+ordered.get(i));
                int width=((Number)p.getPropertyValue("Width")).intValue(),height=((Number)p.getPropertyValue("Height")).intValue();
                require(Math.abs(width-m.widthPt()*2540/72)<=3 && Math.abs(height-m.heightPt()*2540/72)<=3,
                    "Imported shape dimensions/order do not match DOCX for "+m.shapeId());
                properties.add(p);
                originalObjectSizes.add(new int[]{width,height});
            }
            // Same import, untouched geometry. Keep the converter's native result for comparison.
            export(document,Path.of(args[2]));
            // Opt-in for direct Java callers. Only the in-memory compatible
            // export uses explicit fallback; the source and native PDF stay intact.
            applyCjkFontFallback(document,context,args.length==6 ? args[5] : "");
            for(int i=0;i<properties.size();i++) {
                var p=properties.get(i);var size=originalObjectSizes.get(i);
                require(((Number)p.getPropertyValue("Width")).intValue()==size[0]
                    && ((Number)p.getPropertyValue("Height")).intValue()==size[1],
                    "CJK fallback changed an embedded object size");
            }
            StringBuilder audit=new StringBuilder("index\tUNO_name\tshape_id\theight_hmm\ttop_margin_hmm\tbottom_margin_hmm\tw_position_half_pt\tdescent_hmm\tbefore_vertical_orientation\tbefore_position_hmm\tafter_position_hmm\n");
            for(int i=0;i<properties.size();i++){
                var p=properties.get(i);var m=metrics.get(i);
                int height=((Number)p.getPropertyValue("Height")).intValue();
                int width=((Number)p.getPropertyValue("Width")).intValue();
                int descent=(int)Math.round(-m.positionHalfPt()*2540d/144d);
                int topMargin=((Number)p.getPropertyValue("TopMargin")).intValue();
                int bottomMargin=((Number)p.getPropertyValue("BottomMargin")).intValue();
                // Writer adds TopMargin to the baseline anchor BEFORE applying
                // NONE's numeric position. BottomMargin remains line whitespace.
                int position=-height-topMargin+descent;
                Object oldOrientation=p.getPropertyValue("VertOrient"),oldPosition=p.getPropertyValue("VertOrientPosition");
                // Writer's AS_CHARACTER/NONE uses numeric top relative to baseline.
                // Retain the embedded object, its anchor, dimensions and vector preview.
                p.setPropertyValue("VertOrient",VertOrientation.NONE);
                p.setPropertyValue("VertOrientPosition",position);
                int storedPosition=((Number)p.getPropertyValue("VertOrientPosition")).intValue();
                require(Math.abs(storedPosition-position)<=2,"Position update did not stick: requested="+position+", stored="+storedPosition);
                require(((Number)p.getPropertyValue("VertOrient")).shortValue()==VertOrientation.NONE,"Vertical orientation update did not stick");
                require(((Number)p.getPropertyValue("Width")).intValue()==width
                    && ((Number)p.getPropertyValue("Height")).intValue()==height,"Baseline correction changed an object size");
                require(TextContentAnchorType.AS_CHARACTER.equals(p.getPropertyValue("AnchorType")),"Baseline correction changed the inline anchor");
                audit.append(i+1).append('\t').append(ordered.get(i)).append('\t').append(m.shapeId()).append('\t').append(height).append('\t').append(topMargin).append('\t').append(bottomMargin).append('\t')
                    .append(m.positionHalfPt()).append('\t').append(descent).append('\t').append(oldOrientation).append('\t').append(oldPosition).append('\t').append(storedPosition).append('\n');
            }
            require(objects.getElementNames().length==metrics.size(),"Object count changed during correction");
            export(document,Path.of(args[3]));
            Files.writeString(Path.of(args[4]),audit.toString());
            System.out.println("Exported native and baseline-compatible PDFs; mapped every "+metrics.size()+" editable OLE object; source DOCX was not saved");
        } finally { document.dispose(); }
    }
    private static void applyCjkFontFallback(XComponent document,XComponentContext context,String fallback)throws Exception {
        if(fallback.isBlank()) {
            System.out.println("CJK_FONT_FALLBACK disabled");
            return;
        }
        var toolkit=UnoRuntime.queryInterface(com.sun.star.awt.XToolkit.class,
            context.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit",context));
        var device=toolkit==null ? null : toolkit.createScreenCompatibleDevice(1,1);
        if(device==null)throw new FontDependencyException("Cannot enumerate LibreOffice fonts for CJK fallback");
        Set<String> families=new HashSet<>();
        for(var descriptor:device.getFontDescriptors())families.add(descriptor.Name);
        Set<String> installed=normalizeInstalledFonts(families);
        if(installed.isEmpty())throw new FontDependencyException("LibreOffice reported no installed fonts for CJK fallback");
        List<FontPortion> missing=new ArrayList<>();
        Map<XText,String> originalText=new LinkedHashMap<>();
        Set<String> visitedText=new HashSet<>();
        collectMissingCjkFonts(UnoRuntime.queryInterface(XTextDocument.class,document).getText(),installed,missing,originalText,visitedText);
        collectPageStyleCjkFonts(document,installed,missing,originalText,visitedText);
        if(missing.isEmpty()) {
            System.out.println("CJK_FONT_FALLBACK 0 text portions required substitution");
            return;
        }
        fallback=fallback.trim();
        if(fontAvailabilityNormalized(fallback,installed)!=FontAvailability.INSTALLED)throw new FontDependencyException(
            "Configured CJK fallback font is not installed in LibreOffice: "+logFont(fallback)
                +"; "+missing.size()+" CJK text portions request unavailable fonts");
        Map<String,Integer> replaced=new TreeMap<>();
        for(var portion:missing) {
            var properties=portion.properties();
            // Asian fallback must not alter Western/complex-script typography.
            String[] preserved={"CharFontName","CharFontNameComplex","CharHeight","CharHeightAsian","CharHeightComplex"};
            Object[] before=new Object[preserved.length];
            for(int i=0;i<preserved.length;i++)before[i]=properties.getPropertyValue(preserved[i]);
            properties.setPropertyValue("CharFontNameAsian",fallback);
            require(fallback.equals(properties.getPropertyValue("CharFontNameAsian")),"CJK fallback setting did not persist");
            require(portion.text().equals(portion.range().getString()),"CJK fallback changed a text portion");
            for(int i=0;i<preserved.length;i++)require(Objects.equals(before[i],properties.getPropertyValue(preserved[i])),
                "CJK fallback changed unrelated typography: "+preserved[i]);
            replaced.merge(portion.requested(),1,Integer::sum);
        }
        for(var entry:originalText.entrySet())require(entry.getValue().equals(snapshotStaticText(entry.getKey())),
            "CJK fallback changed document text");
        for(var entry:replaced.entrySet())System.out.println("CJK_FONT_FALLBACK "+logFont(entry.getKey())
            +" -> "+logFont(fallback)+"; text portions="+entry.getValue());
        System.out.println("CJK_FONT_FALLBACK total text portions="+missing.size()+"; document text unchanged");
    }
    private static void collectPageStyleCjkFonts(XComponent document,Set<String> installed,List<FontPortion> missing,
                                                 Map<XText,String> originalText,Set<String> visitedText)throws Exception {
        var supplier=UnoRuntime.queryInterface(com.sun.star.style.XStyleFamiliesSupplier.class,document);
        var pageStyles=UnoRuntime.queryInterface(XNameAccess.class,supplier.getStyleFamilies().getByName("PageStyles"));
        for(String name:pageStyles.getElementNames()) {
            Object value=pageStyles.getByName(name);
            var style=UnoRuntime.queryInterface(com.sun.star.style.XStyle.class,value);
            if(style==null || !style.isInUse())continue;
            var properties=UnoRuntime.queryInterface(XPropertySet.class,value);
            for(String property:enabledHeaderFooterProperties(properties)) {
                var text=UnoRuntime.queryInterface(XText.class,properties.getPropertyValue(property));
                if(text!=null)collectMissingCjkFonts(text,installed,missing,originalText,visitedText);
            }
        }
    }
    static List<String> enabledHeaderFooterProperties(XPropertySet properties)throws Exception {
        List<String> result=new ArrayList<>();
        var info=properties.getPropertySetInfo();
        for(String part:new String[]{"Header","Footer"}) {
            if(!info.hasPropertyByName(part+"IsOn") || !Boolean.TRUE.equals(properties.getPropertyValue(part+"IsOn")))continue;
            for(String variant:new String[]{"Text","TextFirst","TextLeft","TextRight"})
                if(info.hasPropertyByName(part+variant))result.add(part+variant);
        }
        return result;
    }
    private static String snapshotStaticText(XText text)throws Exception {
        StringBuilder result=new StringBuilder();
        var elements=UnoRuntime.queryInterface(XEnumerationAccess.class,text).createEnumeration();
        while(elements.hasMoreElements()) {
            var element=elements.nextElement();
            if(UnoRuntime.queryInterface(XTextTable.class,element)!=null)continue; // Cells have their own snapshots.
            var access=UnoRuntime.queryInterface(XEnumerationAccess.class,element);
            if(access==null)continue;
            var portions=access.createEnumeration();
            while(portions.hasMoreElements()) {
                var portion=portions.nextElement();
                var properties=UnoRuntime.queryInterface(XPropertySet.class,portion);
                if(properties==null)continue;
                String type=String.valueOf(properties.getPropertyValue("TextPortionType"));
                if("Text".equals(type))result.append(UnoRuntime.queryInterface(XTextRange.class,portion).getString());
                else if("TextField".equals(type)) {
                    var field=UnoRuntime.queryInterface(XTextField.class,properties.getPropertyValue("TextField"));
                    // PAGE/NUMPAGES displayed values may change after legitimate
                    // repagination; their identities and instructions must not.
                    require(field!=null,"Text field has no field object");
                    result.append('\u0000').append(UnoRuntime.generateOid(field)).append(':')
                        .append(field.getPresentation(true)).append('\u0000');
                }
            }
            result.append('\n');
        }
        return result.toString();
    }
    private static void collectMissingCjkFonts(XText text,Set<String> installed,List<FontPortion> missing,
                                               Map<XText,String> originalText,Set<String> visitedText)throws Exception {
        if(!visitedText.add(UnoRuntime.generateOid(text)))return;
        originalText.put(text,snapshotStaticText(text));
        var enumeration=UnoRuntime.queryInterface(XEnumerationAccess.class,text).createEnumeration();
        while(enumeration.hasMoreElements()) {
            Object element=enumeration.nextElement();
            var table=UnoRuntime.queryInterface(XTextTable.class,element);
            if(table!=null) {
                for(String cell:table.getCellNames())collectMissingCjkFonts(
                    UnoRuntime.queryInterface(XText.class,table.getCellByName(cell)),installed,missing,originalText,visitedText);
                continue;
            }
            var access=UnoRuntime.queryInterface(XEnumerationAccess.class,element);
            if(access==null)continue;
            var portions=access.createEnumeration();
            while(portions.hasMoreElements()) {
                var portion=portions.nextElement();
                var properties=UnoRuntime.queryInterface(XPropertySet.class,portion);
                if(properties==null || !"Text".equals(properties.getPropertyValue("TextPortionType")))continue;
                var range=UnoRuntime.queryInterface(XTextRange.class,portion);
                if(range==null || !containsCjk(range.getString()))continue;
                String requested=String.valueOf(properties.getPropertyValue("CharFontNameAsian"));
                var availability=fontAvailabilityNormalized(requested,installed);
                if(availability==FontAvailability.UNRESOLVED_FAMILY)throw new FontDependencyException(
                    "Cannot safely resolve CJK font family or alias: "+logFont(requested)
                        +"; use a verifiable installed family name or a supported alias");
                if(availability==FontAvailability.MISSING)missing.add(new FontPortion(properties,range,requested,range.getString()));
            }
        }
    }
    static boolean isInstalledFont(String requested,Set<String> installed) {
        return fontAvailability(requested,installed)==FontAvailability.INSTALLED;
    }
    static FontAvailability fontAvailability(String requested,Set<String> installed) {
        return fontAvailabilityNormalized(requested,normalizeInstalledFonts(installed));
    }
    private static Set<String> normalizeInstalledFonts(Set<String> installed) {
        Set<String> names=new HashSet<>();
        for(String family:installed)for(String token:family.split("[;,]")) {
            String normalized=canonicalFontName(token);
            if(!normalized.isEmpty())names.add(normalized);
        }
        return names;
    }
    private static FontAvailability fontAvailabilityNormalized(String requested,Set<String> names) {
        boolean unresolved=false,nonempty=false;
        for(String token:requested.split("[;,]")) {
            String name=canonicalFontName(token);
            if(name.isEmpty())continue;
            nonempty=true;
            if(names.contains(name))return FontAvailability.INSTALLED;
            // A custom ASCII/PostScript alias can be just as ambiguous as a
            // localized alias. Only the finite verified family set is safe to
            // call missing without an installed-font match.
            unresolved|=!FONT_ALIASES.containsValue(name);
        }
        return unresolved || !nonempty ? FontAvailability.UNRESOLVED_FAMILY : FontAvailability.MISSING;
    }
    static String canonicalFontName(String token) {
        // LibreOffice GetNextFontToken accepts comma/semicolon lists. A colon
        // starts OpenType/Graphite feature options, not part of the family name.
        int features=token.indexOf(':');
        if(features>=0)token=token.substring(0,features);
        StringBuilder normalized=new StringBuilder();
        token.codePoints().forEach(value->{
            int c=value;
            if(c>=0xff00 && c<=0xff5e)c-=0xff00-0x20;
            if(c>='A' && c<='Z')c+='a'-'A';
            if(c>127 || c>='a' && c<='z' || c>='0' && c<='9' || c=='(' || c==')')normalized.appendCodePoint(c);
        });
        String name=normalized.toString();
        return FONT_ALIASES.getOrDefault(name,name);
    }
    static boolean containsCjk(String text) {
        return text.codePoints().anyMatch(codePoint -> {
            var script=Character.UnicodeScript.of(codePoint);
            if(script==Character.UnicodeScript.HAN || script==Character.UnicodeScript.HANGUL
                || script==Character.UnicodeScript.HIRAGANA || script==Character.UnicodeScript.KATAKANA
                || script==Character.UnicodeScript.BOPOMOFO)return true;
            // These Common-script characters also use Asian font selection.
            var block=Character.UnicodeBlock.of(codePoint);
            return block==Character.UnicodeBlock.CJK_SYMBOLS_AND_PUNCTUATION
                || block==Character.UnicodeBlock.HALFWIDTH_AND_FULLWIDTH_FORMS
                || block==Character.UnicodeBlock.CJK_COMPATIBILITY_FORMS
                || block==Character.UnicodeBlock.VERTICAL_FORMS
                || block==Character.UnicodeBlock.KATAKANA;
        });
    }
    private static String logFont(String font) {
        return '"'+font.replace("\\","\\\\").replace("\r","\\r").replace("\n","\\n").replace("\t","\\t").replace("\"","\\\"")+'"';
    }
    private static void export(XComponent document,Path path)throws Exception {
        require(!Files.exists(path),"Refusing to overwrite "+path);
        UnoRuntime.queryInterface(XStorable.class,document).storeToURL(path.toAbsolutePath().toUri().toString(),new PropertyValue[]{
            property("FilterName","writer_pdf_Export"),property("Overwrite",false)});
        require(Files.isRegularFile(path)&&Files.size(path)>0,"PDF was not produced: "+path);
    }
    private static List<Metric> readMetrics(Path source)throws Exception {
        var factory=DocumentBuilderFactory.newInstance();factory.setNamespaceAware(true);
        factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl",true);
        factory.setFeature("http://xml.org/sax/features/external-general-entities",false);
        factory.setFeature("http://xml.org/sax/features/external-parameter-entities",false);
        factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD,"");factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_SCHEMA,"");
        factory.setXIncludeAware(false);factory.setExpandEntityReferences(false);
        List<Metric> result=new ArrayList<>();
        try(var zip=new ZipFile(source.toFile());var in=zip.getInputStream(Objects.requireNonNull(zip.getEntry("word/document.xml"),"not a DOCX"))){
            // Do not allow this explicit conversion helper to resolve external
            // package resources. Normal generated exports have internal relations.
            for(var entries=zip.entries();entries.hasMoreElements();){
                var entry=entries.nextElement();if(!entry.getName().endsWith(".rels"))continue;
                try(var relationInput=zip.getInputStream(entry)){
                    var relations=factory.newDocumentBuilder().parse(relationInput).getDocumentElement().getChildNodes();
                    for(int j=0;j<relations.getLength();j++)if(relations.item(j) instanceof Element relation)
                        require(!"External".equals(relation.getAttribute("TargetMode")),"External package relationships are not supported by this helper");
                }
            }
            var xml=factory.newDocumentBuilder().parse(in);
            var directions=xml.getElementsByTagNameNS(W,"textDirection");
            for(int i=0;i<directions.getLength();i++)require("lrTb".equals(((Element)directions.item(i)).getAttributeNS(W,"val")),
                "Only horizontal text layout is supported by this baseline mapping");
            var nodes=xml.getElementsByTagNameNS(W,"object");
            for(int i=0;i<nodes.getLength();i++){
                var object=(Element)nodes.item(i);var ole=(Element)object.getElementsByTagNameNS(O,"OLEObject").item(0);
                require(ole!=null&&"Equation.DSMT4".equals(ole.getAttribute("ProgID"))
                    && "Embed".equals(ole.getAttribute("Type")),"Only embedded MathType Equation.DSMT4 objects are supported");
                var shape=(Element)object.getElementsByTagNameNS(V,"shape").item(0);
                require(shape!=null&&shape.getAttribute("id").equals(ole.getAttribute("ShapeID")),"Missing or mismatched VML shape");
                require(!Pattern.compile("(?:^|;)\\s*(?:rotation|flip)\\s*:").matcher(shape.getAttribute("style")).find(),
                    "Rotated/flipped OLEs need a separate baseline transform");
                var run=(Element)object.getParentNode();require(W.equals(run.getNamespaceURI())&&"r".equals(run.getLocalName()),"Object has no containing run");
                var position=(Element)run.getElementsByTagNameNS(W,"position").item(0);require(position!=null,"Missing DOCX baseline; refusing a guessed offset");
                result.add(new Metric(shape.getAttribute("id"),dimension(shape,"width"),dimension(shape,"height"),Integer.parseInt(position.getAttributeNS(W,"val"))));
            }
        }
        // Plain-text documents have no OLE baseline to correct.
        return result;
    }
    private static double dimension(Element shape,String key){
        var matcher=Pattern.compile("(?:^|;)\\s*"+key+":([0-9]+(?:\\.[0-9]+)?)pt(?:;|$)").matcher(shape.getAttribute("style"));
        require(matcher.find(),"Expected a point-based VML "+key);double value=Double.parseDouble(matcher.group(1));require(value>0&&Double.isFinite(value),"Invalid VML dimension");return value;
    }
    private static void visitText(XText text,Set<String> embeddedNames,List<String> result)throws Exception {
        var enumeration=UnoRuntime.queryInterface(XEnumerationAccess.class,text).createEnumeration();
        while(enumeration.hasMoreElements()){
            Object element=enumeration.nextElement();var table=UnoRuntime.queryInterface(XTextTable.class,element);
            if(table!=null){
                var cells=new ArrayList<>(Arrays.asList(table.getCellNames()));
                for(String cell:cells)require(cell.matches("[A-Z]+[0-9]+"),"Unsupported merged/nested cell name: "+cell);
                cells.sort(Comparator.comparingInt(LibreOfficeBaselineExporter::row).thenComparingInt(LibreOfficeBaselineExporter::column));
                for(String cell:cells)visitText(UnoRuntime.queryInterface(XText.class,table.getCellByName(cell)),embeddedNames,result);
                continue;
            }
            var paragraphs=UnoRuntime.queryInterface(XEnumerationAccess.class,element);if(paragraphs==null)continue;
            var portions=paragraphs.createEnumeration();
            while(portions.hasMoreElements()){
                var portion=portions.nextElement();
                var portionProperties=UnoRuntime.queryInterface(XPropertySet.class,portion);
                if(portionProperties==null || !"Frame".equals(portionProperties.getPropertyValue("TextPortionType")))continue;
                var contents=UnoRuntime.queryInterface(XContentEnumerationAccess.class,portion);if(contents==null)continue;
                var frames=contents.createContentEnumeration("com.sun.star.text.TextContent");
                while(frames.hasMoreElements()){
                    var named=UnoRuntime.queryInterface(XNamed.class,frames.nextElement());
                    if(named!=null&&embeddedNames.contains(named.getName()))result.add(named.getName());
                }
            }
        }
    }
    private static int row(String cell){return Integer.parseInt(cell.replaceAll("[A-Z]",""));}
    private static int column(String cell){int value=0;for(char c:cell.toCharArray()){if(!Character.isLetter(c))break;value=value*26+c-'A'+1;}return value;}
}
