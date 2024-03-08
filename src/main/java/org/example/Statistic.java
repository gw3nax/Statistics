package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;

public class Statistic implements IOManager{

    private final ArrayList<Double> selection = new ArrayList<>();
    private final ArrayList<ArrayList<Double>> relativeFrequencies = new ArrayList<>();
    private final ArrayList<ArrayList<Double>> variationArray = new ArrayList<>();
    private final HashMap<Double, Integer> frequencies = new HashMap<>();
    int n, var;
    long m;
    private String GroupName;
    private String PersonName;
    double minX, maxX, h;
    Statistic(){}

    public void start(){

        Scanner scanner = new Scanner(System.in);

        System.out.print("Введите имя группы: ");
        GroupName = scanner.nextLine();
        System.out.print("Введите ваше имя: ");
        PersonName = scanner.nextLine();
        System.out.print("Введите ваш вариант: ");
        var = scanner.nextInt();
        readData( "./data/" + var + ".txt");

        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph par = doc.createParagraph();
        XWPFRun run = par.createRun();
        run.setFontFamily("Times New Roman");
        run.setFontSize(14);
        run.setText("Работу выполнил: " + PersonName);
        run.addBreak();
        run.setText("Группа: " + GroupName);
        run.addBreak();

        Collections.sort(selection);
        n = selection.size();
        m  = Math.round(1 + 3.322 * Math.log10(n));

        System.out.println("1. Объём выборки = " + n);
        run.setText("1. Объём выборки = " + n);
        run.addBreak();

        minX = CalculateMinElement();
        maxX = CalculateMaxElement();
        System.out.println("2. Наименьшее значение = " + minX + "\n   " +
                "Наибольшее значение = " + maxX);
        run.setText("2. Наименьшее значение = " + minX);
        run.addBreak();
        run.setText("Наибольшее значение = " + maxX);
        run.addBreak();


        System.out.println("3. Размах выборки = " + (maxX - minX));
        run.setText("3. Размах выборки = " + (maxX - minX));
        run.addBreak();


        CalculateFrequencies();
        System.out.println("4. Медиана = " + CalculateMedian() + "\n   " +
                "Мода = " + CalculateModa());
        run.setText("4. Медиана = " + CalculateMedian());
        run.addBreak();
        run.setText("Мода = " + CalculateModa());
        run.addBreak();

        h = (double) Math.round((maxX - minX) * 100 / m) / 100;
        makeVarArray();
        System.out.println("5. Интервальный вариационный ряд:\n   " +
                "Число интервалов = " + m + "\n   " +
                "Длина интервала = " + h + "\n   " +
                "Вариационный ряд:\n   " +
                "Среднее значение | Сумма частот\n   " +
                "--------------------");

        run.setText("5. Интервальный вариационный ряд:");
        run.addBreak();
        run.setText("Число интервалов = " + m);
        run.addBreak();
        run.setText("Длина интервала = " + h);
        run.addBreak();
        run.setText("Вариационный ряд:");
        run.addBreak();

        XWPFTable table = doc.createTable(variationArray.get(0).size(), variationArray.size());

        for (int row = 0; row < variationArray.get(0).size(); row++) {
            for (int col = 0; col < variationArray.size(); col++) {
                XWPFTableCell cell = table.getRow(row).getCell(col);
                cell.setText(variationArray.get(col).get(row).toString());
            }
        }

        for (ArrayList<Double> doubles : variationArray) {
            for (Double aDouble : doubles) {
                System.out.print("   " + aDouble + " ");
            }
            System.out.println();
        }
        System.out.print("""
                   Относительные частоты:
                   Среднее значение | Относительная частота
                   --------------------
                """);

        run.setText("Относительные частоты:");
        run.addBreak();

        XWPFTable table1 = doc.createTable(relativeFrequencies.get(0).size(), relativeFrequencies.size());

        for (int row = 0; row < relativeFrequencies.get(0).size(); row++) {
            for (int col = 0; col < relativeFrequencies.size(); col++) {
                XWPFTableCell cell = table1.getRow(row).getCell(col);
                cell.setText(relativeFrequencies.get(col).get(row).toString());
            }
        }

        for (ArrayList<Double> doubles : relativeFrequencies) {
            for (Double aDouble : doubles) {
                System.out.print("   " + aDouble + " ");
            }
            System.out.println();
        }
        System.out.println("Строим гистограмму относительных частот: ");
        run.setText("Гистограмма относительных частот:");
        run.addBreak();
        run.setText("ВСТАВИТЬ ДИАГРАММУ");
        run.addBreak();

        drawCharts(relativeFrequencies);
        System.out.println("6. Вычисляем точечные оценки параметров распределения: ");
        run.setText("Относительные частоты:");
        run.addBreak();

        double averageValue = CalculateAverageValue();
        double variance = CalculateVariance(averageValue);
        double fixedVariance = variance * ((double)n/(n-1));
        System.out.println("Выборочное среднее = " + averageValue + "\n" +
                "Дисперсия = " + variance + "\n" +
                "Исправленная выборочная дисперсия = " + fixedVariance);

        run.setText("Выборочное среднее = " + averageValue);
        run.addBreak();
        run.setText("Дисперсия = " + variance);
        run.addBreak();
        run.setText("Исправленная выборочная дисперсия = " + fixedVariance);
        run.addBreak();

        double reliability = 0.9;
        double[] StudentArray = {1.664, 1.662, 1.66, 1.658, 1.655, 1.645};
        ArrayList<Double> confidenceInterval = CalculateIntervalEstimation(averageValue, fixedVariance);
        System.out.println("7. Доверительный интервал для мат.ожидания = (" + confidenceInterval.get(0) + " ; " + confidenceInterval.get(1) + ")\n " +
                "Доверительный интервал для среднего квадратичного отклонения = (" + confidenceInterval.get(2) + " ; " + confidenceInterval.get(3) + ")");

        run.setText("7. Доверительный интервал для мат.ожидания = (" + confidenceInterval.get(0) + " ; " + confidenceInterval.get(1) + ")");
        run.addBreak();
        run.setText("Доверительный интервал для среднего квадратичного отклонения = (" + confidenceInterval.get(2) + " ; " + confidenceInterval.get(3) + ")");
        run.addBreak();

        System.out.println("8. Проверяем гипотезу о нормальном распределении: ");
        run.setText("8. Проверяем гипотезу о нормальном распределении: ");
        run.addBreak();
        if (CheckHypothesis(averageValue, fixedVariance)){
            System.out.println("Нет оснований отвергнуть нулевую гипотезу,генеральная совокупность из которой сделана выборка, распределена по нормальному закону");
            run.setText("Нет оснований отвергнуть нулевую гипотезу,генеральная совокупность из которой сделана выборка, распределена по нормальному закону");
            run.addBreak();
        }
        else {
            System.out.println("Распределение ген совокупности не является нормальным");
            run.setText("Распределение ген совокупности не является нормальным");
            run.addBreak();
        }
        try (FileOutputStream fileOut = new FileOutputStream("Отчет.docx")) {
            doc.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                doc.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private double GaussFunc(double u){
        return (1 / Math.sqrt(2 * Math.PI) * Math.exp(-1 * (Math.pow(u,2)/2)));
    }
    private boolean CheckHypothesis(double averageValue, double fixedVariance){
        double Xview = 0, Xcritical = 11.07;
        for (ArrayList<Double> relativeFrequency : variationArray) {
            double u = (relativeFrequency.get(0) - averageValue) / Math.sqrt(fixedVariance);
            double p = h / Math.sqrt(fixedVariance)  * GaussFunc(u);
            double ni = n * p;
            Xview += Math.pow((relativeFrequency.get(1) - ni) , 2) / ni;
        }
        System.out.println(Xview + " | " + Xcritical);
        return Xview < Xcritical;

    }
    private ArrayList<Double> CalculateIntervalEstimation(double averageValue, double fixedVariance){
        double t = 1.655,q = 0.143;
        double leftSideAverage = averageValue - t * (Math.sqrt(fixedVariance) / Math.sqrt(n));
        double rightSideAverage = averageValue + t * (Math.sqrt(fixedVariance)  / Math.sqrt(n));
        double leftSideVariance = q > 1 ? 0 : Math.sqrt(fixedVariance) * (1 - q);
        double rightSideVariance = Math.sqrt(fixedVariance) * (1 + q);
        return new ArrayList<>(Arrays.asList(leftSideAverage, rightSideAverage, leftSideVariance, rightSideVariance));
    }
    private double CalculateVariance(double averageValue){
        double tmp = 0;
        for (ArrayList<Double> doubles : variationArray) {
            tmp += Math.pow((doubles.get(0) - averageValue),2) * doubles.get(1);
        }
        return tmp/n;
    }
    private double CalculateAverageValue(){
        double sum = 0;
        for (ArrayList<Double> doubles : variationArray) {
            sum += doubles.get(0) * doubles.get(1);
        }
        return sum/n;
    }
    private void CalculateFrequencies(){
        for (Double v : selection) {
            frequencies.put(v, frequencies.getOrDefault(v, 0)+1);
        }
    }
    private double CalculateMinElement(){
        return Collections.min(selection);
    }
    private double CalculateMaxElement(){
        return Collections.max(selection);
    }
    private double CalculateMedian(){
        return selection.get(selection.size()/2);
    }
    private double CalculateModa(){
        return Collections.max(frequencies.entrySet(), Map.Entry.comparingByValue()).getKey();
    }
    private void makeVarArray() {
        double currentIntervalStart = minX, currentIntervalEnd = 0;
        int freqSum = 0, cnt = 1;

        Set<Double> unique = new TreeSet<>(selection);
        for (Double value : unique) {
            currentIntervalEnd = currentIntervalStart + h;
            if (value >= currentIntervalStart && value < currentIntervalEnd || cnt == m) {
                freqSum += frequencies.get(value);
            } else {
                double mid = (double) Math.round((currentIntervalStart + currentIntervalEnd) / 2 * 100) / 100;
                relativeFrequencies.add(new ArrayList<>(Arrays.asList(mid, (double) Math.round(((double) freqSum) / n * 100) / 100)));
                variationArray.add(new ArrayList<>(Arrays.asList(mid, (double) freqSum)));
                currentIntervalStart = currentIntervalEnd;
                freqSum = frequencies.get(value);
                cnt++;
            }
        }
        double mid = (double) Math.round((currentIntervalStart + currentIntervalEnd + h) / 2 * 100) / 100;
        variationArray.add(new ArrayList<>(Arrays.asList(mid, (double) freqSum)));
        relativeFrequencies.add(new ArrayList<>(Arrays.asList(mid, (double) Math.round(((double) freqSum) / n * 100) / 100)));
    }
    @Override
    public void readData(String fileName){
        try{
            File file = new File(fileName);
            Scanner scanner = new Scanner(file);
            while (scanner.hasNext()){
                double element = scanner.nextDouble();
                selection.add(element);
            }
            System.out.println("Данные успешно прочитаны.");
            scanner.close();
        } catch (FileNotFoundException e){
            System.err.println("File not found: " + e.getMessage());
        }
    }
    @Override
    public void drawCharts(ArrayList<ArrayList<Double>> arr){
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            XSSFSheet sheet = workbook.createSheet("Histogram");

            int rowNum = 0;
            for (ArrayList<Double> entry : arr) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(entry.get(0));
                row.createCell(1).setCellValue(entry.get(1));
            }

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 5, 1, 20, 15);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Относительные частоты");

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("Значение");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("Частота");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, relativeFrequencies.size(), 0, 0));
            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(0, relativeFrequencies.size(), 1, 1));

            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFChartData.Series series1 = data.addSeries(xs, ys);
            series1.setTitle("Частота", null);
            data.setVaryColors(true);
            chart.plot(data);
            XDDFBarChartData bar = (XDDFBarChartData) data;
            bar.setBarDirection(BarDirection.COL);


            try (FileOutputStream fileOut = new FileOutputStream("histogram.xlsx")) {
                workbook.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            System.out.println("Histogram has been created successfully!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}