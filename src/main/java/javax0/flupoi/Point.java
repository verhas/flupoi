package javax0.flupoi;

class Point {
	static class Coordinate {
		private int value;
		private boolean defined = false;
		private int relativity = 0;

		public int getValue() {
			return value;
		}

		public void setValue(int value) {
			this.value = value;
			this.defined = true;
		}

		public boolean isDefined() {
			return defined;
		}

		public void setDefined(boolean defined) {
			this.defined = defined;
		}

		public int getRelativity() {
			return relativity;
		}

		public void setRelativity(char sig) {
			relativity = sig == '+' ? 1 : (sig == '-' ? -1 : 0);
		}
		public boolean isRelative(){
			return relativity != 0;
		}
	}

	final private Coordinate x = new Coordinate();
	final private Coordinate y = new Coordinate();

	public Coordinate getX() {
		return x;
	}

	public Coordinate getY() {
		return y;
	}

	protected int getXValue() {
		return x.value;
	}

	protected void setXValue(int x) {
		this.x.setValue(x);
	}

	protected int getYValue() {
		return y.value;
	}

	protected void setYValue(int y) {
		this.y.setValue(y);
	}

	protected void setXR(char sig) {
		this.x.setRelativity(sig);
	}

	protected void setYR(char sig) {
		this.y.setRelativity(sig);
	}

	protected boolean xRelative() {
		return this.x.isRelative();
	}

	protected boolean yRelative() {
		return this.y.isRelative();
	}

	protected boolean xDefined() {
		return this.x.defined;
	}

	protected boolean yDefined() {
		return this.y.defined;
	}
}
