package javax0.flupoi;

import java.util.Collection;

import org.apache.poi.ss.usermodel.Cell;

public class EmptyCellMatcherCondition implements Condition {

	@Override
	public boolean match(Collection<Cell> cells) {
		for (Cell cell : cells) {
			return  cell == null || cell.getStringCellValue().length() == 0;
		}
		return true;
	}

}
