package salve;

public class Evento {
	private String week;
	private String type = "";
	private String customer = "";
	private String accountType = "";
	private String description = "";
	private Double effort = 0.0;
	private String partner = "";
	private String preSalePartner = "";
	private String city = "";

	public String getWeek() {
		return week;
	}
	public void setWeek(String week) {
		this.week = week;
	}
	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}
	public String getCustomer() {
		return customer;
	}
	public void setCustomer(String customer) {
		this.customer = customer;
	}
	public String getAccountType() {
		return accountType;
	}
	public void setAccountType(String accountType) {
		this.accountType = accountType;
	}
	public String getDescription() {
		return description;
	}
	public void setDescription(String description) {
		this.description = description;
	}
	public Double getEffort() {
		return effort;
	}
	public void setEffort(Double effort) {
		this.effort = effort;
	}
	public String getPartner() {
		return partner;
	}
	public void setPartner(String partner) {
		this.partner = partner;
	}

	public String getPreSalePartner() {
		return preSalePartner;
	}
	public void setPreSalePartner(String preSalePartner) {
		this.preSalePartner = preSalePartner;
	}
	public String getCity() {
		return city;
	}
	public void setCity(String city) {
		this.city = city;
	}
	@Override
	public String toString() {
		return "Evento [week=" + week + ", type=" + type + ", customer=" + customer + ", accountType=" + accountType
				+ ", description=" + description + ", effort=" + effort + ", partner=" + partner + ", preSalePartner="
				+ preSalePartner + ", city=" + city + "]";
	}

	
}
