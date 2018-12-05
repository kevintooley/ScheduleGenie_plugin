package org.rapla.plugin.schedulegenie;

public class TestShot {
	
	private String Name;
    private String Start;
    private String End;
    private String Resources;
    private String Persons;
    private String duration;
    //private String extraColumn;

    public TestShot(){
    }

    public TestShot(String Name, String Start, String End, String Resources, String Persons, String duration) {
        super();
        this.Name = Name;
        this.Start = Start;
        this.End = End;
        this.Resources = Resources;
        this.Persons = Persons;
        this.duration = duration;
        //this.extraColumn = "";
    }

    //Getters and setters
    public String getName() {
		return Name;
	}

	public void setName(String name) {
		Name = name;
	}

	public String getStart() {
		return Start;
	}

	public void setStart(String start) {
		Start = start;
	}

	public String getEnd() {
		return End;
	}

	public void setEnd(String end) {
		End = end;
	}

	public String getResources() {
		return Resources;
	}

	public void setResources(String resources) {
		Resources = resources;
	}

	public String getPersons() {
		return Persons;
	}

	public void setPersons(String persons) {
		Persons = persons;
	}

	public String getDuration() {
		return duration;
	}

	public void setDuration(String duration) {
		this.duration = duration;
	}

	@Override
    public String toString() {
        return "TestShot [ShotName=" + Name + 
        		", Start=" + Start + 
        		", End=" + End + 
        		", Resources=" + Resources + 
        		", Persons=" + Persons + 
        		", duration=" + duration + "]";
    }

}
