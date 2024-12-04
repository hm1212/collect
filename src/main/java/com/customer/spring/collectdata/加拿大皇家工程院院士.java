package com.customer.spring.collectdata;

import cn.hutool.http.HttpRequest;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import lombok.Data;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.CollectionUtils;

import java.io.*;
import java.text.MessageFormat;
import java.time.Duration;
import java.util.*;

/**
 * <p></p>
 *
 * @author hm 2024/12/2
 * @version V1.0
 * @modificationHistory=========================逻辑或功能性重大变更记录
 * @modify by user: hm 2024/12/2
 * @modify by reason:{方法名}:{原因}
 **/
public class 加拿大皇家工程院院士 {

    private static final Logger LOGGER = LoggerFactory.getLogger(CollectDataApplication.class);

    private static final String PRE_FIX = "https://widgets.cae-acg.ca/";
    private static final String HTML_PREFIX = "F:\\project\\data\\collect\\CollectData\\src\\main\\resources\\static";
    private static final String SHEET_NAME = "加拿大皇家工程院院士.xlsx";

    public static void main(String[] args) {
        // request();
        // webDriver();
        readFile();

    }


    private static void readFile() {
        String path = MessageFormat.format("{0}\\{1}", HTML_PREFIX, "person.html");
        StringBuilder contentBuilder = new StringBuilder();
        try (FileReader reader = new FileReader(path);
             BufferedReader bufferedReader = new BufferedReader(reader)) {
            String line;
            while ((line = bufferedReader.readLine()) != null) {
                contentBuilder.append(line).append(System.lineSeparator());
            }
            String htmlContent = contentBuilder.toString();
            writeFromHtml(htmlContent, HTML_PREFIX, "person.html",false);
        } catch (FileNotFoundException e) {
            System.err.println("文件未找到: " + path);
            e.printStackTrace();
        } catch (IOException e) {
            System.err.println("读取文件时发生错误: " + path);
            e.printStackTrace();
        }
    }

    private static void webDriver() {
        WebDriver driver = null;
        System.setProperty("webdriver.chrome.driver", "F:\\tools\\driver\\win32\\chromedriver.exe");
        // 谷歌驱动
        ChromeOptions options = new ChromeOptions();
        // 允许所有请求
        options.addArguments("--remote-allow-origins=*");
        driver = new ChromeDriver(options);
        driver.get("https://widgets.cae-acg.ca/feeds/directory/directory/action/Alpha/value/ALL/cid/1498/id/2201/listingType/P");
        WebElement iframe = driver.findElement(By.tagName("iframe"));
        driver.switchTo().frame(iframe);
        driver.getPageSource();
        try {
            while (true){
                WebElement loadMore = driver.findElement(By.id("ucDirectory_ucResults_btnLoadMore"));
                loadMore.click();
                String pageSource = driver.getPageSource();
            }
        } catch (NoSuchElementException e) {
            LOGGER.error(e.getMessage());
        } finally {
            if (driver != null) {
                String html = driver.getPageSource();
                writeFromHtml(html, HTML_PREFIX, "person.html",true);
                driver.quit();
            }
        }
    }


    private static void writeFromHtml(String html, String parentPath, String htmlName,boolean isWrite) {
        if (isWrite){
            // 写入文件，作为备份
            writeFile(html, parentPath, htmlName);
        }
        List<List<String>> rows = parseHtml(html);
        // 写出数据
        List<String> header = Arrays.asList("name", "ABOUT ME", "Sector", "Technical Group", "Function");

        writeData(rows, SHEET_NAME, header);
    }

    private static List<List<String>> parseHtml(String html) {
        Document doc = Jsoup.parse(html);
        Elements listings = doc.select(".listing");
        List<Person> list=new ArrayList<>();
        for (Element listing : listings) {
            Element nameElement = listing.select("h3 > div > a").first();
            if (nameElement != null) {
                String name = nameElement.text();
                String href = nameElement.attr("href");
                Person person = new Person();
                person.setName(name);
                person.setUrl(href);
                list.add(person);
            }
        }
        List<List<String>> rows = new ArrayList<>();
        int i = 1;
        if (!CollectionUtils.isEmpty(list)) {
            for (Person person : list) {
                String name = person.getName();
                String hrefValue = person.getUrl();
                String url = MessageFormat.format("{0}{1}", PRE_FIX, hrefValue);
                LOGGER.info("开始获取第:{}个学者数据，本轮总计获取:{}个学者，name:{}, href:{}", i++, list.size(), name, url);
                getPersonInfo(name, url, rows);
            }
        }
        return rows;
    }

    private static void writeFile(String html, String parentPath, String htmlName) {
        FileWriter writer = null;
        try {
            String path = MessageFormat.format("{0}\\{1}", parentPath, htmlName);
            writer = new FileWriter(path);
            writer.write(html);
        } catch (Exception ignore) {
        } finally {
            if (writer != null) {
                try {
                    writer.close();
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
        }
    }

    private static void request() {
        String requestUrl = "https://widgets.cae-acg.ca/feeds/directory/directory/sort/2/action/Alpha/value/ALL/cid/1498/id/2201/listingType/P";
        String html = HttpRequest.get(requestUrl).execute().body();

        // Document doc = Jsoup.parse(html);
        // Elements links = doc.select("a[id=ucDirectory_ucResults_rptResults_ctl00_hlListing]");
        //
        // List<List<String>> rows = new ArrayList<>();
        // for (Element link : links) {
        //     String href = link.attr("href");
        //     if (href != null && !href.isEmpty()) {
        //         getPersonInfo(href, rows);
        //     }
        // }
        Document doc = Jsoup.parse(html);
        Elements listings = doc.select(".listing");

        Map<String, String> persionMap = new HashMap<>();

        for (Element listing : listings) {
            Element nameElement = listing.select("h3 > div > a").first();
            if (nameElement != null) {
                String name = nameElement.text();
                String href = nameElement.attr("href");
                persionMap.put(name, href);
            }
        }
        List<List<String>> rows = new ArrayList<>();
        int i = 0;
        if (!CollectionUtils.isEmpty(persionMap)) {
            for (Map.Entry<String, String> entry : persionMap.entrySet()) {
                String name = entry.getKey();
                String hrefValue = entry.getValue();
                String url = MessageFormat.format("{0}{1}", PRE_FIX, hrefValue);
                LOGGER.info("开始获取第:{}个学者数据，本轮总计获取:{}个学者，name:{}, href:{}", ++i, persionMap.size(), name, url);
                getPersonInfo(name, url, rows);
            }
        }
        // 写出数据
        List<String> header = Arrays.asList("name", "ABOUT ME", "Sector", "Technical Group", "Function");

        writeData(rows, SHEET_NAME, header);
    }

    private static boolean writeData(List<List<String>> rows, String name, List<String> header) {
        ExcelWriter writer = ExcelUtil.getWriter(MessageFormat.format("F:\\project\\data\\collect\\CollectData\\src\\main\\resources\\static\\file\\{0}", name));
        if (!name.endsWith(".xlsx")) {
            name = MessageFormat.format("{0}.xlsx", name);
        }
        writer.setSheet(name);
        writer.writeHeadRow(header);
        writer.write(rows);
        writer.setColumnWidth(-1, 36);
        writer.flush();
        writer.close();
        return true;
    }

    private static void getPersonInfo(String name, String url, List<List<String>> rows) {
        String html = executeCurlCommand(url);
        Document doc = Jsoup.parse(html);
        // 定义一个List，存放每行数据
        List<String> row = new ArrayList<>();
        rows.add(row);
        // name
        //  <div id="ucDirectory_ucListing_divImage" class="logo">
        //                             <img id="ucDirectory_ucListing_imgLogo" class="rounded" title="Abatzoglou, Nicolas" alt="Abatzoglou, Nicolas" src="../../../../../../../../../../../../membershipclientfiles/1498/images/4955697a-bac3-633d-6d1a-2d53de1e02df.jpg" border="0" />
        //                         </div>
        // ABOUT ME
        // <div class="AboutUs">
        //                     <div id="ucDirectory_ucListing_pnlAboutUs" class="ListingSection">
        //
        //                         <h4>About Me</h4>
        //                         <hr>
        //                         <div class="description">
        //         Dr. Abatzoglou has demonstrated achievements with a very high level of impact in pedagogy, academic management and in industry. His main contributions to academia consist of a restructuring of the chemical engineering program and creating the biotech engineering program. As director, the department has seen a dramatic increase in faculty and enrollment. He is recognized as the most popular teacher by students. He developed a very high level research program in collaboration with industry and holds a Chair in pharmaceutical engineering with Pfizer. He is co-founder of two university spin-offs, one of which is Enerkem, and has extensive expertise in technological transfer. He is currently a professor at the University of Sherbrooke.&nbsp;
        //                         </div>
        //
        // </div>
        // <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_divClass" class="div1">
        //                             <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_pnlSection" class="ListingSection">
        //
        //                                 <h4>Other Details</h4>
        //                                 <hr>
        //                                 <div class="other-details">
        //
        //
        //
        //                                             <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_rptFields_ctl01_divField" class="row detail">
        //                                                 <div><strong>Sector:</strong></div>
        //                                                 <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_rptFields_ctl01_divFieldValues"><a href='/feeds/directory/directory/action/Parameter/value/106/cid/1498/id/2201/ListingType/P/Academia'>Academia</a></div>
        //
        //                                             </div>
        //
        //                                             <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_rptFields_ctl02_divField" class="row detail">
        //                                                 <div><strong>Technical Group:</strong></div>
        //                                                 <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_rptFields_ctl02_divFieldValues"><a href='/feeds/directory/directory/action/Parameter/value/107/cid/1498/id/2201/ListingType/P/Chemical-Petrochemical'>Chemical/Petrochemical</a>, <a href='/feeds/directory/directory/action/Parameter/value/452/cid/1498/id/2201/ListingType/P/Education'>Education</a>, <a href='/feeds/directory/directory/action/Parameter/value/127/cid/1498/id/2201/ListingType/P/Environment'>Environment</a></div>
        //
        //                                             </div>
        //
        //                                             <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_rptFields_ctl03_divField" class="row detail">
        //                                                 <div><strong>Function:</strong></div>
        //                                                 <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_rptFields_ctl03_divFieldValues"><a href='/feeds/directory/directory/action/Parameter/value/108/cid/1498/id/2201/ListingType/P/Applied-Research'>Applied Research</a>, <a href='/feeds/directory/directory/action/Parameter/value/144/cid/1498/id/2201/ListingType/P/Design'>Design</a>, <a href='/feeds/directory/directory/action/Parameter/value/109/cid/1498/id/2201/ListingType/P/Teaching'>Teaching</a></div>
        //
        //                                             </div>
        //
        //                                             <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_rptFields_ctl04_divField" class="row detail">
        //                                                 <div><strong>Member of Provincial Engineering Regulator:</strong></div>
        //                                                 <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_rptFields_ctl04_divFieldValues"><a href='/feeds/directory/directory/action/Parameter/value/110/cid/1498/id/2201/ListingType/P/Yes'>Yes</a></div>
        //
        //                                             </div>
        //
        //                                             <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_rptFields_ctl05_divField" class="row detail">
        //                                                 <div><strong>Provincial Engineering Regulator(s):</strong></div>
        //                                                 <div id="ucDirectory_ucListing_rptAdditionalSections_ctl00_rptFields_ctl05_divFieldValues"><a href='/feeds/directory/directory/action/Parameter/value/111/cid/1498/id/2201/ListingType/P/QC'>QC</a></div>
        //
        //                                             </div>
        //
        //                                 </div>
        //
        // </div>
        // Sector
        // Technical Group
        // Function

        // 提取姓名
        if (name == null || name.equalsIgnoreCase("")) {
            name = doc.select("img[alt]").attr("alt");
        }
        row.add(name);

        // 提取 About Me
        String aboutMe = doc.select(".description").text();
        row.add(aboutMe);
        // 提取 Other Details
        Elements details = doc.select(".other-details .row.detail");
        Map<String, String> detailsMap = new HashMap<>();
        for (Element detail : details) {
            String label = detail.select("strong").text();
            Elements values = detail.select("div a");
            StringBuilder valueBuilder = new StringBuilder();
            for (Element value : values) {
                if (valueBuilder.length() > 0) {
                    valueBuilder.append(", ");
                }
                valueBuilder.append(value.text());
            }
            String value = valueBuilder.toString();
            label = label.replace(":", "").trim();
            detailsMap.put(label, value);
        }
        // Sector
        row.add(detailsMap.get("Sector") == null ? "" : detailsMap.get("Sector"));
        // Technical Group
        row.add(detailsMap.get("Technical Group") == null ? "" : detailsMap.get("Technical Group"));
        // Function
        row.add(detailsMap.get("Function") == null ? "" : detailsMap.get("Function"));
    }

    private static String executeCurlCommand(String url) {
        StringBuilder output = new StringBuilder();
        Process p;
        try {
            p = Runtime.getRuntime().exec("curl -g " + url);
            BufferedReader reader = new BufferedReader(new InputStreamReader(p.getInputStream()));
            String line;
            while ((line = reader.readLine()) != null) {
                output.append(line).append("\n");
            }
            reader.close();
            p.waitFor();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return output.toString();
    }
}
@Data
class Person{
    private String name;
    private String url;
}
