import java.util.Set;
import java.util.TreeSet;
import java.util.Map;
import java.util.List;
import java.lang.reflect.Proxy;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;

/** Dependency-light checks; compile alongside the exporter with the public UNO jar. */
public final class LibreOfficeFontFallbackTest {
    private static void check(boolean condition,String message) {
        if(!condition)throw new AssertionError(message);
    }
    public static void main(String[] args)throws Exception {
        for(String text:new String[]{"一、选择题（每题5分，共25分）","，。；（）", "𠀀", "かなカナ", "한글", "ㄅ", "ー"})
            check(LibreOfficeBaselineExporter.containsCjk(text),"CJK text missed: "+text);
        for(String text:new String[]{"", "  \t\n", "English 123", "α + β = ∑x²", "😀"})
            check(!LibreOfficeBaselineExporter.containsCjk(text),"Non-CJK text matched: "+text);
        Set<String> installed=new TreeSet<>(String.CASE_INSENSITIVE_ORDER);
        installed.add("Noto Serif CJK SC");installed.add("User Font");
        check(LibreOfficeBaselineExporter.isInstalledFont("noto serif cjk sc",installed),"Case-insensitive family lookup");
        check(LibreOfficeBaselineExporter.isInstalledFont("Missing; User Font ",installed),"Installed fallback family must be preserved");
        check(LibreOfficeBaselineExporter.isInstalledFont("Missing, User Font ",installed),"Comma-delimited installed family must be preserved");
        check(LibreOfficeBaselineExporter.isInstalledFont("Noto Sans CJK SC;Noto Serif CJK SC:hwid=1",installed),"Font features must not hide an installed family");
        check(LibreOfficeBaselineExporter.isInstalledFont("Ｎｏｔｏ Serif-CJK_SC",installed),"Fullwidth ASCII and font search punctuation normalization");
        check(LibreOfficeBaselineExporter.isInstalledFont("自定义中文字体,User Font",installed),"Unknown first alias must not mask a later installed family");
        check(LibreOfficeBaselineExporter.isInstalledFont("宋体",Set.of("SimSun")),"Chinese SimSun alias");
        check(LibreOfficeBaselineExporter.isInstalledFont("SimSun",Set.of("宋体")),"Localized installed descriptor");
        check(LibreOfficeBaselineExporter.isInstalledFont("楷体_GB2312",Set.of("KaiTi_GB2312")),"Exact legacy KaiTi family alias");
        check(!LibreOfficeBaselineExporter.isInstalledFont("楷体_GB2312",Set.of("KaiTi","SimKai")),"Distinct legacy and modern KaiTi families must not be merged");
        check(LibreOfficeBaselineExporter.isInstalledFont("楷体",Set.of("KaiTi")),"Modern KaiTi alias");
        check(LibreOfficeBaselineExporter.isInstalledFont("ＭＳ ゴシック",Set.of("MS Gothic")),"Localized Japanese alias");
        check(LibreOfficeBaselineExporter.isInstalledFont("바탕",Set.of("Batang")),"Localized Korean alias");
        check(LibreOfficeBaselineExporter.isInstalledFont("任意用户中文字体",Set.of("任意用户中文字体")),"Exact unknown Unicode descriptor must be preserved");
        check(!LibreOfficeBaselineExporter.isInstalledFont("Missing Font",installed),"Missing font classified as installed");
        check(!LibreOfficeBaselineExporter.isInstalledFont(" ; ",installed),"Blank font list classified as installed");
        check(LibreOfficeBaselineExporter.fontAvailability("宋体",installed)==LibreOfficeBaselineExporter.FontAvailability.MISSING,"Known absent font can use fallback");
        check(LibreOfficeBaselineExporter.fontAvailability("任意用户中文字体",installed)==LibreOfficeBaselineExporter.FontAvailability.UNRESOLVED_FAMILY,"Unknown localized alias must fail closed");
        check(LibreOfficeBaselineExporter.fontAvailability("CustomPostScriptAlias",installed)==LibreOfficeBaselineExporter.FontAvailability.UNRESOLVED_FAMILY,"Unknown ASCII alias must fail closed");
        var properties=pageStyle(Map.of("HeaderIsOn",false,"HeaderText","unused",
            "FooterIsOn",true,"FooterText","main","FooterTextFirst","first","FooterTextLeft","left","FooterTextRight","right"));
        check(LibreOfficeBaselineExporter.enabledHeaderFooterProperties(properties).equals(
            List.of("FooterText","FooterTextFirst","FooterTextLeft","FooterTextRight")),"Only enabled footer variants should be traversed");
        check(LibreOfficeBaselineExporter.enabledHeaderFooterProperties(pageStyle(Map.of("HeaderIsOn",true,"HeaderText","main")))
            .equals(List.of("HeaderText")),"Optional unsupported page-style properties must be skipped");
        System.out.println("LibreOffice CJK fallback selection checks passed");
    }
    private static XPropertySet pageStyle(Map<String,Object> values) {
        var info=(XPropertySetInfo)Proxy.newProxyInstance(LibreOfficeFontFallbackTest.class.getClassLoader(),
            new Class<?>[]{XPropertySetInfo.class},(proxy,method,args)->{
                if(method.getName().equals("hasPropertyByName"))return values.containsKey(args[0]);
                throw new UnsupportedOperationException(method.getName());
            });
        return (XPropertySet)Proxy.newProxyInstance(LibreOfficeFontFallbackTest.class.getClassLoader(),
            new Class<?>[]{XPropertySet.class},(proxy,method,args)->{
                if(method.getName().equals("getPropertySetInfo"))return info;
                if(method.getName().equals("getPropertyValue"))return values.get(args[0]);
                throw new UnsupportedOperationException(method.getName());
            });
    }
}
