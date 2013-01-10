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

	private Workbook getSheet() {
		InputStream is = getClass().getClassLoader().getResourceAsStream(
				"test1.xlsx");
		return FluentWorkbookFactory.getWorkbook().stream(is).sheet("Sheet1");
	}

	@Test
	@SuppressWarnings("unchecked")
	public void testAlphaRangeDownPieces() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {

		Workbook wb = getSheet().range("A1:C2").down()
				.names("fruit,animal,travel").target("one", Animal.class)
				.range("(+0,+1):(+2,+0)").down().names("fruit,animal,travel")
				.target("two", Animal.class).execute();
		Collection<Animal> animales = (Collection<Animal>) wb.fetch("one");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
		animales = (Collection<Animal>) (Collection<Animal>) wb.fetch("two");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	private Workbook testExec(Workbook wb) throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		return wb.names("fruit,animal,travel").target("one", Animal.class)
				.execute();
	}

	private Workbook getRange() {
		return getSheet().range("A1:C3");
	}

	@Test
	public void testAlphaRangeDown() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) testExec(
				getRange().down()).fetch("one");
		System.out.println("down");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testAlphaRangeUp() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		InputStream is = getClass().getClassLoader().getResourceAsStream(
				"test1.xlsx");
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A3:C1").up()
				.names("fruit,animal,travel").target("one", Animal.class)
				.execute().fetch("one");
		System.out.println("up");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testAlphaRangeLeft() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		InputStream is = getClass().getClassLoader().getResourceAsStream(
				"test1.xlsx");
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("C3:A1").left()
				.names("fruit,animal,travel").target("one", Animal.class)
				.execute().fetch("one");
		System.out.println("left");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testAlphaRangeRight() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		InputStream is = getClass().getClassLoader().getResourceAsStream(
				"test1.xlsx");
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A1:C3")
				.right().names("fruit,animal,travel")
				.target("one", Animal.class).execute().fetch("one");
		System.out.println("right");
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
				.getWorkbook().stream(is).sheet("Sheet1").range("(0,0):(2,2)")
				.down().names("fruit,animal,travel")
				.target("one", Animal.class).execute().fetch("one");
		System.out.println("num");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}
}
