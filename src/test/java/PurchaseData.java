import java.io.IOException;
import java.util.ArrayList;

public class PurchaseData {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		DataDrivenExcel d =  new DataDrivenExcel();
		
		ArrayList<String> data = d.getDataFromExcel("Purchase");
		System.out.println(data.get(4));

	}

}
