package javax0.flupoi;

import java.io.IOException;
import java.io.InputStream;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

import junit.framework.Assert;
import junit.framework.TestCase;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.junit.experimental.categories.Categories.ExcludeCategory;

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

	private InputStream is;

	@Before
	public void setUp() {
		is = getClass().getClassLoader().getResourceAsStream("test1.xlsx");
	}

	@After
	public void tearDown() {
		try {
			is.close();
		} catch (IOException e) {

		}
	}

	@Test
	@SuppressWarnings("unchecked")
	public void testAlphaRangeDownPieces() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		Workbook wb = FluentWorkbookFactory.getWorkbook().stream(is)
				.sheet("Sheet1").range("A1:C2").down()
				.names("fruit,animal,travel").target("one", Animal.class)
				.range("(+0,+1):(+3,+1)").down().names("fruit,animal,travel")
				.target("two", Animal.class).execute();
		Collection<Animal> animales = (Collection<Animal>) wb.fetch("one");
		System.out.println("pieces");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
		animales = (Collection<Animal>) (Collection<Animal>) wb.fetch("two");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@SuppressWarnings("unchecked")
	@Test
	public void testSkip() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		Workbook wb = FluentWorkbookFactory.getWorkbook().stream(is)
				.sheet("Sheet1").range("A1:C1").down()
				.names("fruit,animal,travel").target("one", Animal.class)
				.skip(0, -1).range("(+0,+0):(+3,+3)").down()
				.names("fruit,animal,travel").target("two", Animal.class)
				.execute();
		System.out.println("pieces backskip");
		Collection<Animal> animales = (Collection<Animal>) wb.fetch("one");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
		animales = (Collection<Animal>) (Collection<Animal>) wb.fetch("two");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testMapTarget() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Map<String, String>> animales = (Collection<Map<String, String>>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A3:C1").down()
				.names("fruit,animal,travel").target("one", HashMap.class)
				.execute().fetch("one");
		System.out.println("map");
		for (Map<String, String> animal : animales) {
			System.out.println(animal.get("animal") + " likes eating "
					+ animal.get("fruit") + " when travelling on a "
					+ animal.get("tracel"));
		}
	}

	@Test
	public void testAlphaRangeDownJoker() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A*:C1").down()
				.names("fruit,animal,travel").target("one", Animal.class)
				.untilEmptyLine().execute().fetch("one");
		System.out.println("down untilEmptyLine");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testAlphaRangeDown() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A3:C1").down()
				.names("fruit,animal,travel").target("one", Animal.class)
				.execute().fetch("one");
		System.out.println("down");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testAlphaRangeDownString() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A1:C*").down()
				.until("bike").names("fruit,animal,travel")
				.target("one", Animal.class).execute().fetch("one");
		System.out.println("down no bike");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testAlphaRangeDownCondition() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A1:C*").down()
				.until(new StringMatcherCondition("bike"))
				.names("fruit,animal,travel").target("one", Animal.class)
				.execute().fetch("one");
		System.out.println("down no bike 2");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testAlphaRangeUpJoker() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A3:C*")
				.names("fruit,animal,travel").target("one", Animal.class)
				.execute().fetch("one");
		System.out.println("up joker");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testAlphaRangeUp() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A3:C1").up()
				.names("fruit,animal,travel").target("one", Animal.class)
				.execute().fetch("one");
		System.out.println("down");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testAlphaRangeLeft() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
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
	public void testAlphaRangeLeftJoker() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("*1:C3")
				.names("fruit,animal,travel").target("one", Animal.class)
				.execute().fetch("one");
		System.out.println("left joker");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testAlphaRangeRightJoker() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A1:*3")
				.untilEmptyCell().names("fruit,animal,travel")
				.target("one", Animal.class).execute().fetch("one");
		System.out.println("right untilEmptyCell");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test
	public void testAlphaRangeRight() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
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
	public void testAlphaRangeNoTarget() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A1:C3")
				.right().names("fruit,animal,travel").execute().fetch("one");
		Assert.assertNull(animales);
	}

	@Test
	public void testAlphaRangeNoName() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A1:C3")
				.right().execute().fetch("one");
		Assert.assertNull(animales);
	}

	@Test
	public void testNumRange() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
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

	@Test
	public void testCell() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").cell("fruit", "A1")
				.target("one", Animal.class).execute().fetch("one");
		System.out.println("cell");
		for (Animal animal : animales) {
			System.out.println(animal);
		}
	}

	@Test(expected = Exception.class)
	public void testException1() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A-:C3")
				.right().execute().fetch("one");
		Assert.assertNull(animales);
	}

	@Test(expected = Exception.class)
	public void testException2() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A1;C3")
				.right().execute().fetch("one");
		Assert.assertNull(animales);
	}

	@Test(expected = Exception.class)
	public void testException3() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("(0;1):(1,3)")
				.right().execute().fetch("one");
		Assert.assertNull(animales);
	}

	@Test(expected = Exception.class)
	public void testException4() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("(0,1)(1,3)")
				.right().execute().fetch("one");
		Assert.assertNull(animales);
	}

	@Test(expected = Exception.class)
	public void testException5() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("(0,1](1,3)")
				.right().execute().fetch("one");
		Assert.assertNull(animales);
	}

	@Test(expected = Exception.class)
	public void testException6() throws InvalidFormatException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException, IOException {
		@SuppressWarnings("unchecked")
		Collection<Animal> animales = (Collection<Animal>) FluentWorkbookFactory
				.getWorkbook().stream(is).sheet("Sheet1").range("A*:*3")
				.right().execute().fetch("one");
		Assert.assertNull(animales);
	}

}
