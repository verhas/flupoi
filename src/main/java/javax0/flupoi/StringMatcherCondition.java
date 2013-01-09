package javax0.flupoi;

import java.util.Collection;

import org.apache.poi.ss.usermodel.Cell;

public class StringMatcherCondition implements Condition {

	private String matcher;

	public StringMatcherCondition(String matcher) {
		this.matcher = matcher;
	}

	@Override
	public boolean match(Collection<Cell> cells) {
		for (Cell cell : cells) {
			String value = cell.getStringCellValue();
			if (value.equalsIgnoreCase(matcher)) {
				return true;
			}
		}
		return false;
	}

}
