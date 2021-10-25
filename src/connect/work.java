package connect;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.*;
import java.util.*;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;

import com.jacob.activeX.*;
import com.jacob.com.*;



public class work {
	//文件處裡 匯出
	/*儲存退出*/
	private boolean saveonexit;
	Dispatch doc = null;
	private ActiveXComponent word;
	private Dispatch documents;
	public static FileSystemView fsv = FileSystemView.getFileSystemView();
	private static JFileChooser filechooser;
	private static FileNameExtensionFilter access = new FileNameExtensionFilter(".accdb", "accdb");
	private static FileNameExtensionFilter docx = new FileNameExtensionFilter(".docx", "docx");
	private static File file;
	static String path = fsv.getHomeDirectory().toString();
	static String filePath = fsv.getHomeDirectory().getAbsolutePath();
	static String datapath = "";
	static String wordpath = "";
	private static String dir_path = path+"/畢業生檔案";
	private static String record_path = path+"/fail_data";
	
	public work()
	{
		if(word==null)
		{
			word = new ActiveXComponent("Word.Application");
			word.setProperty("Visible", new Variant(false));
		}
		if(documents==null)
		{
			documents = word.getProperty("Documents").toDispatch();
		}
		saveonexit = false;
	}
	public boolean setSaveOnExit()
	{
		return saveonexit;
	}
	public Dispatch open(String inputDoc)
	{
		return Dispatch.call(documents, "Open" , inputDoc).toDispatch();
	}
	public Dispatch select() {
        return word.getProperty("Selection").toDispatch();
    }
	public void moveup(Dispatch selection , int count)
	{
		for(int i=0;i<count;i++){
			Dispatch.call(selection, "MoveUp");
		}
	}
	public void movedown(Dispatch selection,int count) {
		for(int i = 0;i < count;i ++) {
			Dispatch.call(selection,"MoveDown");
        }
    }
	public void moveleft(Dispatch selection,int count) {
        for(int i = 0;i < count;i ++) {
            Dispatch.call(selection,"MoveLeft");
        }
    }
	public void moveright(Dispatch selection,int count) {
        for(int i = 0;i < count;i ++) {
            Dispatch.call(selection,"MoveRight");
        }
    }
	public void movestart(Dispatch selection) {
        Dispatch.call(selection,"HomeKey",new Variant(6));
    }
	public boolean find(Dispatch selection,String toFindText) {
        Dispatch find = word.call(selection,"Find").toDispatch();
        Dispatch.put(find,"Text",toFindText);
        Dispatch.put(find,"Forward","True");
        Dispatch.put(find,"Format","True");
        Dispatch.put(find,"MatchCase","True");
        Dispatch.put(find,"MatchWholeWord","True");
        return Dispatch.call(find,"Execute").getBoolean();
    }
	public void replace(Dispatch selection,String newText) {
        Dispatch.put(selection,"Text",newText);
    }
	public void replaceall(Dispatch selection,String oldText,Object replaceObj) {
        movestart(selection);
        if(oldText.startsWith("table") || replaceObj instanceof ArrayList)
            replacetable(selection,oldText,(ArrayList) replaceObj);
        else {
            String newText = (String) replaceObj;
            if(newText==null)
                newText="";
            if(oldText.indexOf("image") != -1&!newText.trim().equals("") || newText.lastIndexOf(".bmp") != -1 || newText.lastIndexOf(".jpg") != -1 || newText.lastIndexOf(".gif") != -1){
            }
            else{
                while(find(selection,oldText)) {
                    replace(selection,newText);
                    Dispatch.call(selection,"MoveRight");
                }
            }
        }
    }
	public void replacetable(Dispatch selection,String tableName,ArrayList dataList) {
        if(dataList.size() <= 1) {
            System.out.println("Empty table!");
            return;
        }
        String[] cols = (String[]) dataList.get(0);
        String tbIndex = tableName.substring(tableName.lastIndexOf("@") + 1);
        int fromRow = Integer.parseInt(tableName.substring(tableName.lastIndexOf("$") + 1,tableName.lastIndexOf("@")));
        Dispatch tables = Dispatch.get(doc,"Tables").toDispatch();
        Dispatch table = Dispatch.call(tables,"Item",new Variant(tbIndex)).toDispatch();
        Dispatch rows = Dispatch.get(table,"Rows").toDispatch();
        for(int i = 1;i < dataList.size();i ++) {
            String[] datas = (String[]) dataList.get(i);
            if(Dispatch.get(rows,"Count").getInt() < fromRow + i - 1) {
            	Dispatch.call(rows,"Add");
            }
            for(int j = 0;j < datas.length;j++) {
                Dispatch cell = Dispatch.call(table,"Cell",Integer.toString(fromRow + i - 1),cols[j]).toDispatch();
                Dispatch.call(cell,"Select");
                Dispatch font = Dispatch.get(selection,"Font").toDispatch();
                Dispatch.put(font,"Bold","0");
                Dispatch.put(font,"Italic","0");
                Dispatch.put(selection,"Text",datas[j]);
            }
        }
	}
	public void save(String outputPath) {
        Dispatch.call(Dispatch.call(word,"WordBasic").getDispatch(),"FileSaveAs",outputPath);
    }
	public void close(Dispatch doc) {
        Dispatch.call(doc,"Close",new Variant(saveonexit));
        word.invoke("Quit",new Variant[]{});
        word = null;
    }
	public void toWord(String inputPath,String outPath,HashMap data) {
       String oldText;
       Object newValue;
       try {
            if(doc==null)
            doc = open(inputPath);
            Dispatch selection = select();
            Iterator keys = data.keySet().iterator();
            while(keys.hasNext()) {
            	oldText = (String) keys.next();
                newValue = data.get(oldText);
                replaceall(selection,oldText,newValue);
            }
             save(outPath);
       } catch(Exception e) {
            System.out.println("Use word fail");
            e.printStackTrace();
       } finally {
            if(doc != null)
            	close(doc);
       }
	}
	
	public synchronized static void word(String inputPath,String outPath,HashMap data){
        work j2w = new work();
        j2w.toWord(inputPath,outPath,data);
    }
	
	
	public static void pushword(ArrayList worklist , int[] sum , int StudentID_Amount,String[] str_GraduationYear,String[] str_Semester,String[] str_Class,String[] str_StudentID,int i)
	{
		int Amount;
		int start = 0, end = 0;
		Amount = StudentID_Amount/4;
		if(i!=0)
		{
			end = i*Amount+Amount;
		}
		if(i==3)
		{
			end = StudentID_Amount;
		}
		start = i*Amount;
		int k=0,x=0;
		HashMap data = new HashMap();
		ArrayList table1 = new ArrayList(7);
		String[] fieldtest = {"1","2","3","4","5","6","7"};
		ComThread.InitSTA();
		for(i=start;i<end;i++)
		{
			table1.add(fieldtest);
			x=0;
			String[] take_record = (String[]) worklist.get(i);
			while(x<sum[i])
			{
				if(x==0)
				{
					data.put("$number$", take_record[x]);
					x++;
				}
				else if(x==1)
				{
					data.put("$class$", take_record[x]);
					x++;
				}
				else if(x==2)
				{
					data.put("$year$", take_record[x]);
					x++;
				}
				else if((sum[i]-x)==7)
				{
					data.put("$sum1$",take_record[x]);
					x++;
				}
				else if((sum[i]-x)==6)
				{
					data.put("$sum2$",take_record[x]);
					x++;
				}
				else if((sum[i]-x)==5)
				{
					data.put("$sum3$",take_record[x]);
					x++;
				}
				else if((sum[i]-x)==4)
				{
					data.put("$sum4$",take_record[x]);
					x++;
				}
				else if((sum[i]-x)==3)
				{
					data.put("$sum12$",take_record[x]);
					x++;
				}
				else if((sum[i]-x)==2)
				{
					data.put("$sum34$",take_record[x]);
					x++;
				}
				else if((sum[i]-x)==1)
				{
					data.put("$sumall$",take_record[x]);
					x++;
				}
				else
				{
					String[] field = {take_record[x],take_record[x+1],take_record[x+2],take_record[x+3],take_record[x+4],take_record[x+5],take_record[x+6]};
					table1.add(field);
					data.put("table$4@1",table1);
					x+=7;
				}
			}
			
			work jw2 = new work();
			jw2.toWord(wordpath,dir_path+"\\"+str_GraduationYear[i]+str_Semester[i]+"_"+str_Class[i]+"_"+str_StudentID[i]+".doc", data);
			table1.clear();
		}
		ComThread.Release();
	}
	
	static class MyThread extends Thread
	{
		int i = 0;
		int Amount;
		ArrayList list = new ArrayList();
		int[] sum;
		int StudentID_Amount;
		String str_GraduationYear[];
		String str_Semester[];
		String str_Class[];
		String str_StudentID[];
		public MyThread(ArrayList worklist , int[] sum , int StudentID_Amount,String[] str_GraduationYear,String[] str_Semester,String[] str_Class,String[] str_StudentID,int i) {
			this.list = worklist;
			this.sum = sum;
			this.StudentID_Amount = StudentID_Amount;
			this.str_GraduationYear = str_GraduationYear;
			this.str_Semester = str_Semester;
			this.str_Class = str_Class;
			this.str_StudentID = str_StudentID;
			this.i = i;
		}
		public void run()
		{
			pushword(list,sum,StudentID_Amount,str_GraduationYear,str_Semester,str_Class,str_StudentID,i);
		}
	}
	
	
	
	static int MAX = 1500;
	public static void main(String[] args) throws SQLException, IOException
	{		
		Connection connDB = null;
		
		ArrayList<String[]> list = new ArrayList<String[]>();
		
		String str_TakeCourse_Number;	// 從修課資料的學號
		String str_TakeCourse_Name;	// 存修課資料的課程名稱
		int int_TakeCourse_Semester;			// 存修課資料的修課學期
		String str_TakeCourse_Grade;	// 存修課資料的修課年級
		String str_TakeCourse_Type;	// 存修課資料的必選修
		String str_TakeCourse_SemesterType;	// 存修課資料的上下學期
		
		String str_Semester[] = new String[MAX];
		String StudentID = new String();
		
		String str_StudentID[] = new String[MAX];	// 存基本資料的學號
		String str_Class[] = new String[MAX];	// 存基本資料的組別
		String str_GraduationYear[] = new String[MAX];	// 存基本資料的畢業學年
		int int_Semester;  // 存基本資料的畢業學期
		
	    String str_Course_Name;	// 存課程的課程名稱
		int int_Course_Credit[] = new int[4];		// 存課程的學分
		
		int StudentID_Amount = 0;
		
		// 各個學生的資料分配
		String str_Student_SemesterType[] = new String[8];
		String str_Student_Name[][] = new String[8][100];
		String str_Student_Type[][] = new String[8][100];
		int int_Student_Credit[][][] = new int[8][100][4];
		// 專題研究與實作的資料存放
		String str_Student_SpectialSemesterType[] = new String[2];
		String str_Student_SpectialName[] = new String[2];
		String str_Student_SpectialType[] = new String[2];
		int int_Student_SpectialCredit[][] = new int[2][4];
		
		long time1=0,time2=0;
		String fail;
		Path p = Paths.get(dir_path);
		if(!Files.exists(p))
		{
			try {
				Files.createDirectory(p);
			}catch (IOException e1){
				e1.printStackTrace();
			}
		}
		p = Paths.get(record_path);
		if(!Files.exists(p))
		{
			try {
				Files.createDirectory(p);
			}catch (IOException e1){
				e1.printStackTrace();
			}
		}
		File writename1 = new File(record_path+"\\工程.txt");
		File writename2 = new File(record_path+"\\數科.txt");
		File writename3 = new File(record_path+"\\基礎科學.txt");
		File writename4 = new File(record_path+"\\數學.txt");
		
		BufferedWriter out1 = new BufferedWriter(new FileWriter(writename1));
		BufferedWriter out2 = new BufferedWriter(new FileWriter(writename2));
		BufferedWriter out3 = new BufferedWriter(new FileWriter(writename3));
		BufferedWriter out4 = new BufferedWriter(new FileWriter(writename4));
		
		do {
			work.filechooser = new JFileChooser(path);
			work.filechooser.setMultiSelectionEnabled(true);
			work.filechooser.setDialogTitle("請選擇資料庫檔案");
			work.filechooser.setFileFilter(access);
			int returnVal = work.filechooser.showOpenDialog(null);
            if (returnVal == 0) {
            	work.filePath = work.filechooser.getSelectedFile().getAbsolutePath();
            	work.file = work.filechooser.getSelectedFile();
           	    datapath = work.file.getAbsolutePath().toString();
            }
		}while(datapath.toString().equals(""));
		
		do {
			work.filechooser = new JFileChooser(path);
			work.filechooser.setMultiSelectionEnabled(true);
			work.filechooser.setDialogTitle("請選擇Word檔案");
			work.filechooser.setFileFilter(docx);
			int returnVal = work.filechooser.showOpenDialog(null);
			if (returnVal == 0) {
				work.filePath = work.filechooser.getSelectedFile().getAbsolutePath();
				work.file = work.filechooser.getSelectedFile();
				wordpath = work.file.getAbsolutePath().toString();
			}
		}while(wordpath.toString().equals(""));
		
		try
		{
			Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
			String path = "jdbc:ucanaccess://"+datapath;
			connDB = DriverManager.getConnection(path);
			Statement st1 = connDB.createStatement();
			Statement st2 = connDB.createStatement();
			Statement st3 = connDB.createStatement();
			ResultSet rs_Information;
			ResultSet rs_Pass;
			ResultSet rs_Course;
			int summ = 0;
			int sums = 0;
			int sumt = 0;
			int sumd = 0;
			int sum1 = 0;
			int sum2 = 0;
			int sumall = 0;
			
			String mathc = new String();
			String sciencec = new String();
			String theoremc = new String();
			String designc= new String();
			String sum12 = new String();
			String sum34 = new String();
			String all = new String();
			
			String combine = new String();
			String math = new String();
			String science = new String();
			String theorem = new String();
			String design= new String();
			int z=0,k=0;
			int a[] = new int[8];
			int Course_math,Course_science,Course_theorem,Course_design;
			int passone = 0;
			int passtwo = 0;
			int sum[]= new int[MAX];
			
			for(int i=0;i<MAX;i++)
			{
				sum[i] = 0;
			}
			
			rs_Information = st1.executeQuery("select 學號,畢業學系,畢業學年,畢業學期  from 基本資料");
			while(rs_Information.next())
			{
				String[] record = new String[1500];
				str_StudentID[k] = rs_Information.getString("學號").toString();
				str_Class[k] = rs_Information.getString("畢業學系");
				str_GraduationYear[k] = rs_Information.getString("畢業學年");
				int_Semester = rs_Information.getInt("畢業學期");
				if(int_Semester==1)
				{
					str_Semester[k]="上";
				}
				else
				{
					str_Semester[k]="下";
				}
				if(str_StudentID[k].length()==6)
				{
					StudentID = str_StudentID[k].substring(3);
					record[sum[k]] = StudentID;
					sum[k]++;
				}
				else if(str_StudentID[k].length()==7)
				{
					StudentID = str_StudentID[k].substring(4);
					record[sum[k]] = StudentID;
					sum[k]++;
				}
				else if(str_StudentID[k].length()==9)
				{
					StudentID = str_StudentID[k].substring(6);
					record[sum[k]] = StudentID;
					sum[k]++;
				}
				if(str_Class[k].length()==11)
				{
					str_Class[k] = str_Class[k].substring(6);
					record[sum[k]] = str_Class[k];
					sum[k]++;
				}
				else
				{
					record[sum[k]] = str_Class[k];
					sum[k]++;
				}
				record[sum[k]] = str_GraduationYear[k];
				sum[k]++;
				for(z=0;z<8;z++)
				{
					a[z]=0;
				}
				rs_Pass = st2.executeQuery("select 修課時就讀年級,課程名稱,修課學期,必選修 from 修課資料_及格  where 學號 ='"+ str_StudentID[k] + "'");
				while(rs_Pass.next())
				{
					str_TakeCourse_Name=rs_Pass.getString("課程名稱").toString();
					int_TakeCourse_Semester=rs_Pass.getInt("修課學期");
					str_TakeCourse_Grade=rs_Pass.getString("修課時就讀年級");
					str_TakeCourse_Type=rs_Pass.getString("必選修");
					if(int_TakeCourse_Semester==1)
					{
						str_TakeCourse_SemesterType="上";
					}
					else
					{
						str_TakeCourse_SemesterType="下";
					}
					rs_Course = st3.executeQuery("select 數學,基礎科學,專業理論,專業設計  from 課程 where 課程名稱 ='"+ str_TakeCourse_Name +"'");
					if (rs_Course.next()) {
			            // read the data out of the result set.
						int_Course_Credit[0] = rs_Course.getInt("數學");
						int_Course_Credit[1] = rs_Course.getInt("基礎科學");
						int_Course_Credit[2] = rs_Course.getInt("專業理論");
						int_Course_Credit[3] = rs_Course.getInt("專業設計");
			        } 
							combine= str_TakeCourse_Grade+str_TakeCourse_SemesterType;
							if(str_TakeCourse_Name.equals("專題研究"))
							{
								str_Student_SpectialSemesterType[0] = combine;
								str_Student_SpectialName[0]=str_TakeCourse_Name;
								str_Student_SpectialType[0]=str_TakeCourse_Type;
								int_Student_SpectialCredit[0][0]=int_Course_Credit[0];
								int_Student_SpectialCredit[0][1]=int_Course_Credit[1];
								int_Student_SpectialCredit[0][2]=int_Course_Credit[2];
								int_Student_SpectialCredit[0][3]=int_Course_Credit[3];
							passone=1;
							}
							else if(str_TakeCourse_Name.equals("專題實作"))
							{
								str_Student_SpectialSemesterType[1] = combine;
								str_Student_SpectialName[1]=str_TakeCourse_Name;
								str_Student_SpectialType[1]=str_TakeCourse_Type;
								int_Student_SpectialCredit[1][0]=int_Course_Credit[0];
								int_Student_SpectialCredit[1][1]=int_Course_Credit[1];
								int_Student_SpectialCredit[1][2]=int_Course_Credit[2];
								int_Student_SpectialCredit[1][3]=int_Course_Credit[3];
								passtwo=1;
							}
							else if(str_TakeCourse_Grade.equals("1"))
							{
								if(int_TakeCourse_Semester==1)
								{
							
									str_Student_SemesterType[0] = combine;
									str_Student_Type[0][a[0]] = str_TakeCourse_Type;
									str_Student_Name[0][a[0]] = str_TakeCourse_Name;
									Course_math = int_Course_Credit[0];
									Course_science = int_Course_Credit[1];
									Course_theorem = int_Course_Credit[2];
									Course_design = int_Course_Credit[3];
									int_Student_Credit[0][a[0]][0]=Course_math;
									int_Student_Credit[0][a[0]][1]=Course_science;
									int_Student_Credit[0][a[0]][2]=Course_theorem;
									int_Student_Credit[0][a[0]][3]=Course_design;
									a[0]++;
								}
								else
								{
									str_Student_SemesterType[1] = combine;
									str_Student_Name[1][a[1]] = str_TakeCourse_Name;
									str_Student_Type[1][a[1]] = str_TakeCourse_Type;
									Course_math = int_Course_Credit[0];
									Course_science = int_Course_Credit[1];
									Course_theorem = int_Course_Credit[2];
									Course_design = int_Course_Credit[3];
									int_Student_Credit[1][a[1]][0]=Course_math;
									int_Student_Credit[1][a[1]][1]=Course_science;
									int_Student_Credit[1][a[1]][2]=Course_theorem;
									int_Student_Credit[1][a[1]][3]=Course_design;
									a[1]++;
								}
							}
							else if(str_TakeCourse_Grade.equals("2"))
							{
								if(int_TakeCourse_Semester==1)
								{
									str_Student_SemesterType[2] = combine;
									str_Student_Name[2][a[2]] = str_TakeCourse_Name;
									str_Student_Type[2][a[2]] = str_TakeCourse_Type;
									Course_math = int_Course_Credit[0];
									Course_science = int_Course_Credit[1];
									Course_theorem = int_Course_Credit[2];
									Course_design = int_Course_Credit[3];
									int_Student_Credit[2][a[2]][0]=Course_math;
									int_Student_Credit[2][a[2]][1]=Course_science;
									int_Student_Credit[2][a[2]][2]=Course_theorem;
									int_Student_Credit[2][a[2]][3]=Course_design;
									a[2]++;
								}
								else
								{
									str_Student_SemesterType[3] = combine;
									str_Student_Name[3][a[3]] = str_TakeCourse_Name;
									str_Student_Type[3][a[3]] = str_TakeCourse_Type;
									Course_math = int_Course_Credit[0];
									Course_science = int_Course_Credit[1];
									Course_theorem = int_Course_Credit[2];
									Course_design = int_Course_Credit[3];
									int_Student_Credit[3][a[3]][0]=Course_math;
									int_Student_Credit[3][a[3]][1]=Course_science;
									int_Student_Credit[3][a[3]][2]=Course_theorem;
									int_Student_Credit[3][a[3]][3]=Course_design;
									a[3]++;
								}
							}
							else if(str_TakeCourse_Grade.equals("3"))
							{
								if(int_TakeCourse_Semester==1)
								{
									str_Student_SemesterType[4] = combine;
									str_Student_Name[4][a[4]] = str_TakeCourse_Name;
									str_Student_Type[4][a[4]] = str_TakeCourse_Type;
									Course_math = int_Course_Credit[0];
									Course_science = int_Course_Credit[1];
									Course_theorem = int_Course_Credit[2];
									Course_design = int_Course_Credit[3];
									int_Student_Credit[4][a[4]][0]=Course_math;
									int_Student_Credit[4][a[4]][1]=Course_science;
									int_Student_Credit[4][a[4]][2]=Course_theorem;
									int_Student_Credit[4][a[4]][3]=Course_design;
									a[4]++;
								}
								else
								{
									str_Student_SemesterType[5] = combine;
									str_Student_Name[5][a[5]] = str_TakeCourse_Name;
									str_Student_Type[5][a[5]] = str_TakeCourse_Type;
									Course_math = int_Course_Credit[0];
									Course_science = int_Course_Credit[1];
									Course_theorem = int_Course_Credit[2];
									Course_design = int_Course_Credit[3];
									int_Student_Credit[5][a[5]][0]=Course_math;
									int_Student_Credit[5][a[5]][1]=Course_science;
									int_Student_Credit[5][a[5]][2]=Course_theorem;
									int_Student_Credit[5][a[5]][3]=Course_design;
									a[5]++;
								}
							}
							else if(str_TakeCourse_Grade.equals("4"))
							{
								if(int_TakeCourse_Semester==1)
								{
									str_Student_SemesterType[6] = combine;
									str_Student_Name[6][a[6]] = str_TakeCourse_Name;
									str_Student_Type[6][a[6]] = str_TakeCourse_Type;
									Course_math = int_Course_Credit[0];
									Course_science = int_Course_Credit[1];
									Course_theorem = int_Course_Credit[2];
									Course_design = int_Course_Credit[3];
									int_Student_Credit[6][a[6]][0]=Course_math;
									int_Student_Credit[6][a[6]][1]=Course_science;
									int_Student_Credit[6][a[6]][2]=Course_theorem;
									int_Student_Credit[6][a[6]][3]=Course_design;
									a[6]++;
								}
								else
								{
									str_Student_SemesterType[7] = combine;
									str_Student_Name[7][a[7]] = str_TakeCourse_Name;
									str_Student_Type[7][a[7]] = str_TakeCourse_Type;
									Course_math = int_Course_Credit[0];
									Course_science = int_Course_Credit[1];
									Course_theorem = int_Course_Credit[2];
									Course_design = int_Course_Credit[3];
									int_Student_Credit[7][a[7]][0]=Course_math;
									int_Student_Credit[7][a[7]][1]=Course_science;
									int_Student_Credit[7][a[7]][2]=Course_theorem;
									int_Student_Credit[7][a[7]][3]=Course_design;
									a[7]++;
								}
							}
					summ+=int_Course_Credit[0];
					sums+=int_Course_Credit[1];
					sumt+=int_Course_Credit[2];
					sumd+=int_Course_Credit[3];
				}
				int x=0,b=0;
				for(x=0;x<8;x++)
				{
					for(b=0;b<100;b++)
					{
						if(b==a[x])
						{
							break;
						}
						else
						{
							math = Integer.toString(int_Student_Credit[x][b][0]);
							science = Integer.toString(int_Student_Credit[x][b][1]);
							theorem = Integer.toString(int_Student_Credit[x][b][2]);
							design = Integer.toString(int_Student_Credit[x][b][3]);
							record[sum[k]] = str_Student_SemesterType[x];
							sum[k]++;
							record[sum[k]]  = str_Student_Name[x][b];
							sum[k]++;
							record[sum[k]] = str_Student_Type[x][b];
							sum[k]++;
							record[sum[k]] = math;
							sum[k]++;
							record[sum[k]] = science;
							sum[k]++;
							record[sum[k]] = theorem;
							sum[k]++;
							record[sum[k]] = design;
							sum[k]++;
						}
					}
					if(x==7)
					{
						if(passone==1)
						{
							math = Integer.toString(int_Student_SpectialCredit[0][0]);
							science = Integer.toString(int_Student_SpectialCredit[0][1]);
							theorem = Integer.toString(int_Student_SpectialCredit[0][2]);
							design = Integer.toString(int_Student_SpectialCredit[0][3]);
							
							record[sum[k]] = str_Student_SpectialSemesterType[0];
							sum[k]++;
							record[sum[k]]  = str_Student_SpectialName[0];
							sum[k]++;
							record[sum[k]] = str_Student_SpectialType[0];
							sum[k]++;
							record[sum[k]] = math;
							sum[k]++;
							record[sum[k]] = science;
							sum[k]++;
							record[sum[k]] = theorem;
							sum[k]++;
							record[sum[k]] = design;
							sum[k]++;
						}
						if(passtwo==1)
						{
							math = Integer.toString(int_Student_SpectialCredit[1][0]);
							science = Integer.toString(int_Student_SpectialCredit[1][1]);
							theorem = Integer.toString(int_Student_SpectialCredit[1][2]);
							design = Integer.toString(int_Student_SpectialCredit[1][3]);
							
							record[sum[k]] = str_Student_SpectialSemesterType[1];
							sum[k]++;
							record[sum[k]]  = str_Student_SpectialName[1];
							sum[k]++;
							record[sum[k]] = str_Student_SpectialType[1];
							sum[k]++;
							record[sum[k]] = math;
							sum[k]++;
							record[sum[k]] = science;
							sum[k]++;
							record[sum[k]] = theorem;
							sum[k]++;
							record[sum[k]] = design;
							sum[k]++;
						}
					}
				}
				
				
				
				
				
				mathc = Integer.toString(summ);
				record[sum[k]] = mathc;
				sum[k]++;
				sciencec = Integer.toString(sums);
				record[sum[k]] = sciencec;
				sum[k]++;
				theoremc = Integer.toString(sumt);
				record[sum[k]] = theoremc;
				sum[k]++;
				designc = Integer.toString(sumd);
				record[sum[k]] = designc;
				sum[k]++;
				
				sum1 = summ+sums;
				sum2 = sumt+sumd;
				sumall = sum1+sum2;
				
				sum12 = Integer.toString(sum1);
				record[sum[k]] = sum12;
				sum[k]++;
				sum34 = Integer.toString(sum2);
				record[sum[k]] = sum34;
				sum[k]++;
				all = Integer.toString(sumall);
				record[sum[k]] = all;
				sum[k]++;
				
				
				list.add(record);
				if(sum2<48) {
					try {
						writename1.createNewFile(); // 建立新檔案
						fail = str_StudentID[k]+"\r\n";
						out1.write(fail); // \r\n即為換行
					}catch (Exception e) {
						//System.out.print("工程");
						e.printStackTrace();
					}
				}
				else if(sum1<32) {
					try {
						writename2.createNewFile(); // 建立新檔案
						fail = str_StudentID[k]+"\r\n";
						out2.write(fail); // \r\n即為換行
					}catch (Exception e) {
						//System.out.print("數學&基礎科學");
						e.printStackTrace();
					}
				}
				else if(sums<9) {
					try {
						writename3.createNewFile(); // 建立新檔案
						fail = str_StudentID[k]+"\r\n";
						out3.write(fail); // \r\n即為換行
					}catch (Exception e) {
						//System.out.print("基礎科學");
						e.printStackTrace();
					}
				}
				else if(summ<9) {
					try 
					{
						writename4.createNewFile(); // 建立新檔案
						fail = str_StudentID[k]+"\r\n";
						out4.write(fail); // \r\n即為換行
					}catch (Exception e) {
						//System.out.print("數學");
						e.printStackTrace();
					}
				}
				
				summ=0;
				sums=0;
				sumt=0;
				sumd=0;
				sum1=0;
				sum2=0;
				sumall=0;
				k++;
				StudentID_Amount++;
			}
			Thread t1 = new MyThread(list,sum,StudentID_Amount,str_GraduationYear,str_Semester,str_Class,str_StudentID,0);
			Thread t2 = new MyThread(list,sum,StudentID_Amount,str_GraduationYear,str_Semester,str_Class,str_StudentID,1);
			Thread t3 = new MyThread(list,sum,StudentID_Amount,str_GraduationYear,str_Semester,str_Class,str_StudentID,2);
			Thread t4 = new MyThread(list,sum,StudentID_Amount,str_GraduationYear,str_Semester,str_Class,str_StudentID,3);
			t1.start();
			t2.start();
			t3.start();
			t4.start();
		}catch(ClassNotFoundException e)
		{
			
		}
		out1.flush(); // 把快取區內容壓入檔案
		out1.close(); 
		out2.flush(); // 把快取區內容壓入檔案
		out2.close(); 
		out3.flush(); // 把快取區內容壓入檔案
		out3.close(); 
		out4.flush(); // 把快取區內容壓入檔案
		out4.close(); 
	}
} 