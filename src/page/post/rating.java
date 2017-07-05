package page.post;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.restfb.Connection;
import com.restfb.DefaultFacebookClient;
import com.restfb.FacebookClient;
import com.restfb.Parameter;
import com.restfb.types.OpenGraphRating;
import com.restfb.types.Page;
import com.restfb.types.PageRating;

public class rating {
	public static void main(String[] args) throws IOException {
    	//access token generatied from developers.facebook.com graph api explorer using an app
        String Accesstoken="EAAbRhKFCvKUBAMZCxhWrw3b6Ogg5ZC4vUFio0zHZBdWPxbYHdkNNFTH8qAIDNNvj46lzGFrXrRhLsMSxZBZAUZAg8lNBGy8DMLW4fi0h8ZBBbj7uQ1sm8rFbxVsuhlyrZC0yNtWmL0KZAmXZBKjeLvrhKgtnZAZCLZCQxSZCkZD";
        
        //initializing a facebook object using restfb library
		@SuppressWarnings("deprecation")
		FacebookClient Fb=new DefaultFacebookClient(Accesstoken);
        Page page =Fb.fetchObject("Dance.Society.DTU", Page.class);
        
        HSSFWorkbook wb1 = new HSSFWorkbook();
		
		HSSFSheet p_rating = wb1.createSheet();
		wb1.setSheetName(0, "Rating");
		
		
		//initialize column headers for page rating.
		 HSSFRow r5 = p_rating.createRow(0);
	    HSSFCell rt1 = r5.createCell(0);
		HSSFCell rt2 = r5.createCell(1);
		HSSFCell rt3 = r5.createCell(2);
		HSSFCell rt4 = r5.createCell(3);
		HSSFCell rt5 = r5.createCell(4);
		HSSFCell rt6 = r5.createCell(5);
			
		rt1.setCellValue("Reviewer's Name");
		rt2.setCellValue("Date and Time");
		rt3.setCellValue("Rating Scale");
		rt4.setCellValue("Rating Value");
		rt5.setCellValue("Rating Id");
		rt6.setCellValue("Review Text");
		
		 System.out.println( page.getId());
	        System.out.println(page.getName());
		
		
		  //code for Rating 
		int rating_row=1;
           Connection<OpenGraphRating> reviewnew  = Fb.fetchConnection(page.getId()+"/ratings", OpenGraphRating.class,Parameter.with("fields", "open_graph_story{id,message,publish_time,from{name,id},data{rating,language,review_text}}"));
  			for (List<OpenGraphRating> reviewpage :reviewnew) {
                   for (OpenGraphRating rev : reviewpage) {
                	  
                      PageRating rating = rev.getOpenGraphStory();
                      
                      //create new row
                      HSSFRow r6 = p_rating.createRow(rating_row);
                      
                      	//initialize cell from where to start storing
						int cell=0;
						HSSFCell rc1 = r6.createCell(cell);
						HSSFCell rc2 = r6.createCell(cell+1);
						HSSFCell rc3=  r6.createCell(cell+2);
						HSSFCell rc4 = r6.createCell(cell+3);
						HSSFCell rc5 = r6.createCell(cell+4);
						HSSFCell rc6 = r6.createCell(cell+5);
						
						//set cell width
						p_rating.setColumnWidth(0, 10000);
						p_rating.setColumnWidth(1, 8000);
						p_rating.setColumnWidth(2, 5000);
						p_rating.setColumnWidth(3, 5000);
						p_rating.setColumnWidth(4, 10000);
						p_rating.setColumnWidth(5, 50000);
                    
						//extract and save name of person who rated the page.
                      System.out.println(rating.getFrom().getName()); 
                      rc1.setCellValue(new HSSFRichTextString(rating.getFrom().getName())) ;
                      
                      //extract and strore the publish date and time of the rating.
                      System.out.println("Create time :"+rating.getPublishTime());
                      Date dt2= new Date();
                      dt2=rating.getPublishTime();
                      DateFormat df2=DateFormat.getDateTimeInstance();
                      rc2.setCellValue(new HSSFRichTextString(df2.format(dt2))) ;
                      
                      //extract and store the rating scale print the Rating.
                      System.out.println("Rating Scale:"+rating.getRatingScale());
                      rc3.setCellValue(rating.getRatingScale() );
                      
                      //extract and store rating value given by reviewer
                      System.out.println("Rating Value:"+rating.getRatingValue());
                      rc4.setCellValue(rating.getRatingValue()) ;
                      
                      //extract abd store rating id
                      System.out.println("Rating id:"+rating.getId());
                      rc5.setCellValue(rating.getId()) ;
                      
                      //extract and store review text
                      System.out.println(rating.getReviewText());
                      rc6.setCellValue(new HSSFRichTextString(rating.getReviewText())) ;
                      
                      rating_row++;
                          
                   	} 
  			}
  			FileOutputStream out1 = new FileOutputStream("Rating.xls");
			wb1.write(out1);
			out1.close();
			
			System.out.println("Program run successful. Data is stored in Rating.xls");
	}
}

