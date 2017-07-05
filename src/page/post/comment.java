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
import org.apache.poi.ss.usermodel.CellStyle;


import com.restfb.Connection;
import com.restfb.DefaultFacebookClient;
import com.restfb.FacebookClient;
import com.restfb.types.Comment;
import com.restfb.types.Page;
import com.restfb.types.Post;

public class comment {

	  public static void main(String[] args) throws IOException {
	    	//access token generatied from developers.facebook.com graph api explorer using an app
	        String Accesstoken="EAAbRhKFCvKUBAMZCxhWrw3b6Ogg5ZC4vUFio0zHZBdWPxbYHdkNNFTH8qAIDNNvj46lzGFrXrRhLsMSxZBZAUZAg8lNBGy8DMLW4fi0h8ZBBbj7uQ1sm8rFbxVsuhlyrZC0yNtWmL0KZAmXZBKjeLvrhKgtnZAZCLZCQxSZCkZD";
	        
	        //initializing a facebook object using restfb library
			@SuppressWarnings("deprecation")
			FacebookClient Fb=new DefaultFacebookClient(Accesstoken);
	        Page page =Fb.fetchObject("Dance.Society.DTU", Page.class);
	        HSSFWorkbook wb = new HSSFWorkbook();
	        
			// create 3 new sheet (post,comment and pageratingand name them
			HSSFSheet  c_post= wb.createSheet();
			wb.setSheetName(0, "comment");
			
			//initialize column headers for post comments.
			HSSFRow r4 = c_post.createRow(0);
	    	HSSFCell c1 = r4.createCell(0);
			HSSFCell c2 = r4.createCell(1);
			HSSFCell c3= r4.createCell(2);
			HSSFCell c4= r4.createCell(3);
			HSSFCell c5= r4.createCell(4);
			c1.setCellValue("Post ID");
			c2.setCellValue("Comment ID");
			c3.setCellValue("Person who commented");
			c4.setCellValue("Date and Time");
			c5.setCellValue("Message");
			//fetch comment of post
			
			 System.out.println( page.getId());
		        System.out.println(page.getName());
			
			int com_row=1;
			
			Connection<Post> postfeed = Fb.fetchConnection(page.getId()+"/feed",Post.class);
			for(List<Post> postPage  : postfeed   ){
    	  				for(Post apost : postPage )
					{
			Connection<Comment> commentConnection  = Fb.fetchConnection(apost.getId() + "/comments", Comment.class);
				for (List<Comment> commentPage : commentConnection) {
						for (Comment comment : commentPage) {
							//HSSFSheet worksheet1 = workbook.getSheet("Comment");
							
							
						//	int rowCCount = worksheet1.getLastRowNum();
								HSSFRow r3 = c_post.createRow(com_row);	//create new row
								
								//initialize cell from where to start storing
								int cell=0;									
								HSSFCell cc1 = r3.createCell(cell);
								HSSFCell cc2 = r3.createCell(cell+1);
								HSSFCell cc3= r3.createCell(cell+2);
								HSSFCell cc4 = r3.createCell(cell+3);
								HSSFCell cc5 = r3.createCell(cell+4);
								
								//set cell width
								c_post.setColumnWidth(0, 10000);
						        c_post.setColumnWidth(1, 10000);
						        c_post.setColumnWidth(2, 8000);
						        c_post.setColumnWidth(3, 10000);
						        c_post.setColumnWidth(4, 54000);
						        
						        System.out.println(com_row);
						        CellStyle st = wb.createCellStyle(); //Create new style
					            st.setWrapText(true); //Set wordwrap
					            
					            //store post id
					            cc1.setCellValue(apost.getId());
							
								//extract and store comment id
								//	System.out.println("Id  "+c+" :" + comment.getId());
								cc2.setCellValue(comment.getId());
                    
								//extract and save the name of person who commented
								//  System.out.println(comment.getFrom().getName());
								cc3.setCellValue(new HSSFRichTextString(comment.getFrom().getName())) ;
                    
								//store the creation date and time after converting to string
								Date dt1= new Date();
								dt1=comment.getCreatedTime();
								DateFormat df1=DateFormat.getDateTimeInstance();
								cc4.setCellValue(new HSSFRichTextString(df1.format(dt1))) ;
								
								//extract and store message of comment
								// System.out.println("Message :"+comment.getMessage());
								cc5.setCellStyle(st);
								cc5.setCellValue(new HSSFRichTextString(comment.getMessage())) ;
                    
								com_row++; //update comment row
								}
							}
					
						}
    	  			}
			FileOutputStream out = new FileOutputStream("comments.xls");
			wb.write(out);
			out.close();
			
			
			System.out.println("Program run successful. Data is stored in comments.xls");
	  	}
}	
