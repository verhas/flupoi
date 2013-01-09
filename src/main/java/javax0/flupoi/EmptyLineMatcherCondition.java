package javax0.flupoi;

import java.util.Collection;

import org.apache.poi.ss.usermodel.Cell;

public class EmptyLineMatcherCondition implements Condition {

	@Override
	public boolean match(Collection<Cell> cells) {
		for (Cell cell : cells) {
			if (cell != null && cell.getStringCellValue().length() > 0) {
				return false;
			}
		}
		return true;
	}

}
