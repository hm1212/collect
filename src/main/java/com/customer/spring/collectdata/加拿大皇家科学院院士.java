package com.customer.spring.collectdata;

import com.customer.spring.util.FileUtil;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class 加拿大皇家科学院院士 {
    private static final Logger LOGGER = LoggerFactory.getLogger(加拿大皇家科学院院士.class);
    private static final List<List<String>> ROWS = new ArrayList<>();
    private static final String SHEET_NAME = "加拿大皇家科学院院士.xlsx";
    private static final List<String> HEADER = Arrays.asList("Name", "Affiliation", "Academy or College", "Year Elected", "Website URL");


    private static void webDriver() throws InterruptedException {
        WebDriver driver = null;

        System.setProperty("webdriver.chrome.driver", "F:\\tools\\driver\\win32\\chromedriver.exe");
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");
        driver = new ChromeDriver(options);

        beforeClickParam(driver);
        WebElement currentPageElement = driver.findElement(By.cssSelector(".pager-current"));
        String html = driver.getPageSource();
        getPersonInfo(html, ROWS);
        clickPageNext(driver, currentPageElement.getText());
        if (!ROWS.isEmpty()) {
            FileUtil.writeData(ROWS, SHEET_NAME, HEADER);
        }
        driver.quit();
    }

    private static void clickPageNext(WebDriver driver, String pageNum){
        int targetPage = Integer.parseInt(pageNum) + 1;
        if (targetPage == 218) {
            return;
        }
        try {
            List<WebElement> elements = driver.findElements(By.className("pager-item"));
            for (WebElement element : elements) {
                WebElement a = element.findElement(By.tagName("a"));
                String title = a.getAttribute("title");
                if (title.equals("Go to page " + targetPage)) {
                    a.click();
                    waitForPageLoad(driver);
                    String html = driver.getPageSource();
                    getPersonInfo(html, ROWS);
                    clickPageNext(driver, String.valueOf(targetPage));
                    return;
                }
            }
        } catch (UnhandledAlertException e) {
            LOGGER.error("正向获取中断，中断页码:{},尝试从后往前获取", targetPage);
            driver.quit();
            ChromeOptions options = new ChromeOptions();
            driver = new ChromeDriver(options);
            beforeClickParam(driver);
            redoClick(String.valueOf(targetPage - 1), driver);
        } catch (StaleElementReferenceException | NoSuchElementException e) {
            LOGGER.info("pageNum：{}", targetPage);
        }
    }

    private static void clickPagePre(WebDriver driver, String pageNum, String endPageNum) throws InterruptedException {
        int targetPage = Integer.parseInt(pageNum) - 1;
        if (targetPage == Integer.parseInt(endPageNum)) {
            return;
        }
        List<WebElement> elements = driver.findElements(By.className("pager-item"));
        for (WebElement element : elements) {
            WebElement a = element.findElement(By.tagName("a"));
            String title = a.getAttribute("title");
            if (title.equals("Go to page " + targetPage)) {
                a.click();
                waitForPageLoad(driver);
                String html = driver.getPageSource();
                getPersonInfo(html, ROWS);
                clickPagePre(driver, String.valueOf(targetPage), endPageNum);
                return;
            }
        }
    }

    public static void redoClick(String startPage, WebDriver driver) {
        beforeClickParam(driver);
        WebElement lastPageElement = driver.findElement(By.className("pager-last"));
        lastPageElement.click();
        waitForPageLoad(driver);
        WebElement currentPageElement = driver.findElement(By.className(".pager-current"));
        int currentPage = Integer.parseInt(currentPageElement.getText()) + 1;
        try {
            clickPagePre(driver, String.valueOf(currentPage), startPage);
        } catch (InterruptedException e) {
            LOGGER.error("中断异常", e);
        }
    }

    private static void getPersonInfo(String html, List<List<String>> rows) {
        Document doc = Jsoup.parse(html);
        Elements rowsElements = doc.select(".views-row");
        for (Element row : rowsElements) {
            List<String> list = new ArrayList<>();
            Elements nameElements = row.select(".views-field-display-name h3.field-content");
            String name = nameElements.size() > 0 ? nameElements.get(0).text() : "";
            list.add(name);

            Elements affiliationElements = row.select(".views-field-current-employer-1 span.field-content");
            String affiliation = affiliationElements.size() > 0 ? affiliationElements.get(0).text() : "";
            list.add(affiliation);

            Elements academyElements = row.select(".views-field-academy-25 span.field-content");
            String academy = academyElements.size() > 0 ? academyElements.get(0).text() : "";
            list.add(academy);

            Elements yearElectedElements = row.select(".views-field-election-year-21 span.field-content");
            String yearElected = yearElectedElements.size() > 0 ? yearElectedElements.get(0).text() : "";
            list.add(yearElected);

            Elements websiteElements = row.select(".views-field-url span a");
            String websiteUrl = websiteElements.size() > 0 ? websiteElements.get(0).attr("href") : "";
            list.add(websiteUrl);
            rows.add(list);
        }
    }

    private static void beforeClickParam(WebDriver driver) {
        driver.get("https://rsc-src.ca/en/find-rsc-member/results?combine=&first_name=&last_name=&current_employer=&academy_25=Academy+of+the+Arts+and+Humanities&is_deceased=All");
        Select select = new Select(driver.findElement(By.id("edit-academy-25")));
        select.selectByValue("All");
        WebElement searchMembers = driver.findElement(By.id("edit-submit-search-fellows"));
        searchMembers.click();
        waitForPageLoad(driver);
        System.out.println("请求成功");
    }

    private static void waitForPageLoad(WebDriver driver) {
        try {
            Thread.sleep(2000);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
    }

    public static void main(String[] args) {
        try {
            webDriver();
        } catch (InterruptedException e) {
            LOGGER.error("主线程中断", e);
        }
    }
}
