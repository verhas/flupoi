package javax0.flupoi;

import java.io.IOException;
import java.io.InputStream;
import java.util.Collection;

import javax0.flupoi.FluentWorkbookFactory;

import junit.framework.TestCase;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

/**
 * Unit test for simple App.
 */
public class TestWorkbook extends TestCase {

	public static class Animal {
		public String fruit;
		public String animal;
		public String travel;

		public String toString() {
			return animal + " likes eating " + fruit + " when travelling on a "
					+ travel;
		}
	}

	@Test
	public void testAlphaRange() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		InputStream is = getClass().getClassLoader().getResourceAsStream(
				"test1.xlsx");
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A1:C3").down()
				.names("fruit,animal,travel").target("one", Animal.class)
				.execute().fetch("one");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}
	@Test
	public void testNumRange() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		InputStream is = getClass().getClassLoader().getResourceAsStream(
				"test1.xlsx");
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("(0,0):(2,2)").down()
				.names("fruit,animal,travel").target("one", Animal.class)
				.execute().fetch("one");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}
}
