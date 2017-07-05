package page.post;

import java.io.FileInputStream;
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
import org.apache.poi.ss.usermodel.CellStyle;

import com.restfb.Connection;
import com.restfb.DefaultFacebookClient;
import com.restfb.FacebookClient;
import com.restfb.Parameter;
import com.restfb.types.Page;
import com.restfb.types.Post;

public class up {
    public static void main(String[] args) throws IOException {
    	//access token generatied from developers.facebook.com graph api explorer using an app
        String Accesstoken="EAAbRhKFCvKUBAMZCxhWrw3b6Ogg5ZC4vUFio0zHZBdWPxbYHdkNNFTH8qAIDNNvj46lzGFrXrRhLsMSxZBZAUZAg8lNBGy8DMLW4fi0h8ZBBbj7uQ1sm8rFbxVsuhlyrZC0yNtWmL0KZAmXZBKjeLvrhKgtnZAZCLZCQxSZCkZD";
        
        //initializing a facebook object using restfb library
		@SuppressWarnings("deprecation")
		FacebookClient Fb=new DefaultFacebookClient(Accesstoken);
        Page page =Fb.fetchObject("Dance.Society.DTU", Page.class) ;
        FileInputStream fileInputStream = new FileInputStream("workbook1.xls");
	HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
		
       
        //create a workbook
        HSSFWorkbook wb = new HSSFWorkbook();
        
		// create 3 new sheet (post,comment and pageratingand name them
		HSSFSheet fb_post = wb.createSheet();
		
		
		wb.setSheetName(0, "Post");		
		//get pageid and name
        System.out.println( page.getId());
        System.out.println(page.getName());
        
       //initialize column headers for page post.
        HSSFRow r1 = fb_post.createRow(0);
    	HSSFCell p1 = r1.createCell(0);
		HSSFCell p2 = r1.createCell(1);
		HSSFCell p3 = r1.createCell(2);
		HSSFCell p4 = r1.createCell(3);
		HSSFCell p5 = r1.createCell(4);
		
		p1.setCellValue("Post ID");
		p2.setCellValue("Posted By");
		p3.setCellValue("Poster's ID");
		p4.setCellValue("Date and Time");
		p5.setCellValue("Message");
		
		
		
       // HSSFSheet worksheet = workbook.getSheet("Post");
	//	HSSFRow row1 = worksheet.getRow(1);
		//HSSFCell cellA1 = row1.getCell((short) 0);
		//String a1Val = cellA1.getStringCellValue();
	//	System.out.println(a1Val);
	//	int rowCount = worksheet.getLastRowNum();
		//int b=rowCount;
     // 	System.out.println(rowCount);	
	//	initialize rows from where post and comment will store.	
		
      	int post_row=1;
		
       
		/*fetch post from facebook page "Dance Society dtu"
		 * using external restfb library availiable from developers.facebook.com
		 *  using connection<post> and post.class from restfb for post 
		 *  and parameter tag to retrieve various fields 
		 */
		Connection<Post> postfeed = Fb.fetchConnection(page.getId()+"/feed",Post.class, Parameter.with("fields", "id,created_time,message,from"));
			for(List<Post> postPage  : postfeed   ){
    	  
					for(Post apost : postPage )
					{  
						
						
			//				int a= apost.getId().compareTo(a1Val);
				//			System.out.println(a);
	              //  	   if (a==0){break;
	                //	   }
	                	//   else{
						
						//create new row
						HSSFRow r2 = fb_post.createRow(post_row);
						
						//initialize column from where to start storing
						int cellnum=0;
						HSSFCell pc1 = r2.createCell(cellnum);
						HSSFCell pc2 = r2.createCell(cellnum+1);
						HSSFCell pc3 = r2.createCell(cellnum+2);
						HSSFCell pc4 = r2.createCell(cellnum+3);
						HSSFCell pc5 = r2.createCell(cellnum+4);
						
						//set width of the columns
						fb_post.setColumnWidth(0, 10000);
				        fb_post.setColumnWidth(1, 6000);
				        fb_post.setColumnWidth(2, 6000);
				        fb_post.setColumnWidth(3, 8000);
				        fb_post.setColumnWidth(4, 54000);
				        
				        //Create new style
				        CellStyle style = wb.createCellStyle(); 
			            style.setWrapText(true); //Set wordwrap
			          
			         
			            //extract post id and store it Print the post ID.
						System.out.println("post "+post_row+" :"+apost.getId());
						pc1.setCellValue(apost.getId());
						
						//extract name of person who posted
						//System.out.println(apost.getFrom().getName());
				        pc2.setCellValue(apost.getFrom().getName());
				        
				        //extract id of person who posted
				        //System.out.println(apost.getFrom().getId());
				        pc3.setCellValue(apost.getFrom().getId());
    
				        //extract creation date and time convert to string format
						//System.out.println("$$$$"+apost.getCreatedTime());
				        Date dt= new Date();
						dt= apost.getCreatedTime();
						DateFormat df=DateFormat.getDateInstance();
						pc4.setCellValue(new HSSFRichTextString(df.format(dt))) ;
						
						//extract the message of the post.
						//System.out.println("---->"+apost.getMessage());
						pc5.setCellValue(new HSSFRichTextString(apost.getMessage())) ;   
						pc5.setCellStyle(style);  //Apply style to cell
			
						post_row++;  //update post row
					
                           // }
				
	              	   
					}
					//break;
				
	  		//save the output into xls file	
			//if(rowCount!=b)
			//{
			
		
			//}			
			
       	}
			FileOutputStream out = new FileOutputStream("workbook1.xls");
			wb.write(out);	
			out.close();
			//fileInputStream.close();
			System.out.println("Program run successful. Data is stored in workbook1.xls");
    }}
