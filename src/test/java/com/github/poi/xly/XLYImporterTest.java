package com.github.poi.xly;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.util.List;

import org.apache.commons.beanutils.Converter;
import org.junit.Before;
import org.junit.Test;

import com.github.poi.xly.test.TestBananas;
import com.github.poi.xly.test.TestScenario;
import com.github.poi.xly.test.TestWorkbook;
import com.github.poi.xly.test.XLYFactory;

public class XLYImporterTest {

    public XLYImporter<TestWorkbook> xlyImporter;
    private final SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");

    @Before
    public void setup() {
        XLYFactory.setup();
    }

    /**
     * Test IllegalArgumentException if null workbook object.
     */
    @Test(expected = IllegalArgumentException.class)
    public void testSave_nullWorkbook() {
        xlyImporter = new XLYImporter<>();
        xlyImporter.setWorkbookClass(null);
        final InputStream inputStream = XLYFactory.getBananasOK();
        xlyImporter.save(inputStream);
    }

    /**
     * Test IllegalArgumentException if null inputStream object.
     */
    @Test(expected = IllegalArgumentException.class)
    public void testSave_nullInputstream() {
        xlyImporter = new XLYImporter<>();
        xlyImporter.setWorkbookClass(TestWorkbook.class);
        xlyImporter.save(null);
    }

    @Test
    public void testSave() {
        xlyImporter = new XLYImporter<>();
        xlyImporter.setWorkbookClass(TestWorkbook.class);
        final InputStream inputStream = XLYFactory.getBananasOK();
        final TestWorkbook workbook = xlyImporter.save(inputStream);
        assertNotNull(workbook);
        assertFlowns(workbook.getBananas());
        assertScenario(workbook.getScenarios());
    }

    @Test
    public void testGetWorkbookClass() {
        xlyImporter = new XLYImporter<>();
        xlyImporter.setWorkbookClass(TestWorkbook.class);
        assertEquals(TestWorkbook.class, xlyImporter.getWorkbookClass());

    }

    @Test
    public void testRegister() {
        xlyImporter = new XLYImporter<>();
        xlyImporter.setWorkbookClass(TestWorkbook.class);
        xlyImporter.register(new CustomConverter(), DayOfWeek.class);
        // ConvertUtilsBean#register should work as expected... not testing it.
    }

    private class CustomConverter implements Converter {
        @SuppressWarnings("unchecked")
        @Override
        public <T> T convert(Class<T> type, Object value) {
            return (T) DayOfWeek.valueOf(value.toString());
        }

    }

    private void assertScenario(final List<TestScenario> scenario) {
        assertNotNull(scenario);
        assertEquals(1, scenario.size());
        assertEquals("SCENARIO", scenario.get(0).getName());
    }

    private void assertFlowns(final List<TestBananas> flowns) {
        assertNotNull(flowns);
        assertEquals(2, flowns.size());
        final TestBananas franceToAlgeria = flowns.get(0);
        assertEquals("FR", franceToAlgeria.getAgencies());
        assertEquals("08/06/2018", formatter.format(franceToAlgeria.getCreationDate()));
        assertEquals(2, franceToAlgeria.getQuantity().intValue());
        assertEquals(2, franceToAlgeria.getRevenue().intValue());
        assertEquals("email1@corp1.com", franceToAlgeria.getEmail());
        assertEquals("FR", franceToAlgeria.getOrigin());
        assertEquals("AL", franceToAlgeria.getDestination());

        final TestBananas algeriaToFrance = flowns.get(1);
        assertEquals("FR", algeriaToFrance.getAgencies());
        assertEquals("08/06/2018", formatter.format(algeriaToFrance.getCreationDate()));
        assertEquals(2, algeriaToFrance.getQuantity().intValue());
        assertEquals(2, algeriaToFrance.getRevenue().intValue());
        assertEquals("email2@corp2.com", algeriaToFrance.getEmail());
        assertEquals("AL", algeriaToFrance.getOrigin());
        assertEquals("FR", algeriaToFrance.getDestination());
    }
}
