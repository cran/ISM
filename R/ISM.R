#'Interpretive Structural Modeling (ISM).
#'
#'This methods provides a wellformated solution of ISM
#'
#'@return provides two output files (Final Reachability Matrix and Level Partition of each iteration)
#'in Excel format
#'@author Adarsh Anand, Gunjan Bansal
#'
#'@details This Function Provides well-formatted
#'and readable excel output files (Final Reachability Matrix and Level Partition of each iteration) that
#'make interpretation easier.
#'
#'@param fname a matrix consists of 1s' and 0's (initial reachability matrix)
#'@param Dir a path where user wants to save output files
#'@import("xlsx","rJava","xlsxjars")
#'
#'@examples ISM(fname=matrix(c(1,1,1,1,1,0,1,1,1,1,0,0,1,0,0,0,1,1,1,1,0,1,1,0,1),5,5,byrow=TRUE),Dir=tempdir())
#'
#'@references Adarsh Anand, Gunjan Bansal, (2017) "Interpretive structural modelling for attributes of software quality", Journal of Advances in Management Research, Vol. 14 Issue: 3, pp.256-269, https://doi.org/10.1108/JAMR-11-2016-0097
#'
#'
#'@export
#'

ISM = function(fname, Dir)
{
  A_mat <- as.data.frame(fname)

  clenght<-as.numeric(dim(A_mat)[2])

  cn <- vector("numeric",clenght)
  for(i in 1:clenght)
  {
    cn[i]<- paste0("A",i)
  }

  colnames(A_mat) <- cn

  len<- length(A_mat)




  flag<-1
  b_row<-c("","","","","")
  final_mat<-matrix(nrow=0,ncol=5)
  Heading<-c("Variable_Names","Reachability_Set","Antecedents_Set","Intersection_Set","Level")


  while (len>0)
  {
    B_mat <- t(A_mat)
    new_matrix <- A_mat
    i=1
    j=1
    k=1

    col<-colnames(A_mat)
    rown<-rownames(A_mat)<-col

    r=nrow(A_mat)
    c=ncol(A_mat)


    for(i in 1:r)
    {
      for (j in 1:c)
      {
        if (i==j)
        {
          new_matrix[i,j]=1

        }
        else if (A_mat[i,j]==1)
        {
          for(k in 1:c)
          {
            if (j==k)
            {
              new_matrix[j,k]=1
            }
            else if (A_mat[j,k]==1 && A_mat[i,k]==0)
            {
              new_matrix[i,k]=1
            }
          }
        }
      }
    }



    fin_mat=new_matrix
    trans_nwmat=t(new_matrix)

    if (flag==1)
    {
      wb2<-xlsx::createWorkbook(type = "xlsx")
      wbn1<-"ISM_Matrix"
      sheet2 <- xlsx::createSheet(wb2, sheetName="New_Matrix")
      xlsx::addDataFrame(x=fin_mat, sheet=sheet2, characterNA="",row.names=FALSE)
      xlsx::autoSizeColumn(sheet2, colIndex=1:c)
      rows  <- xlsx::getRows(sheet2)
      xlsx::setRowHeight(rows, inPoints=20)
      file2 <- paste(wbn1,".xlsx", sep="")

      p<-gsub(" ","",(paste(Dir,"/",file2)))

      print(p)
      xlsx::saveWorkbook(wb2,p)
      options(warn=-1)
      Mat_format(fin_mat,A_mat,p)



      flag=2

    }

    reachset<- vector("list",r)
    antiset <-  vector("list",r)
    interset<- vector("list",r) #intersect(reachset,antiset)

    x<-names(A_mat)

    for (i in 1:r)
    {
      for (j in 1:c)
      {
        if (fin_mat[i,j]==1)
          reachset[[i]]=cbind(reachset[[i]],x[j])
      }
    }

    for (i in 1:r)
    {
      for (j in 1:c)
      {

        if (trans_nwmat[i,j]==1)
          antiset[[i]]=cbind(antiset[[i]],x[j])
      }

    }

    for (i in 1:r)
    {
      interset[[i]]<- intersect(reachset[[i]],antiset[[i]])
    }


    levelout<- vector("numeric",r)


    for (i in 1:r)
    {
      if (length(interset[[i]])==length(reachset[[i]]))
      {  if (all.equal(as.vector(interset[[i]]),as.vector(reachset[[i]])))
      {
        levelout[i]=1
      }
      }
      else
      {
        levelout[i]=0
      }
    }


    ###

    final1<-matrix(nrow=r,ncol=5)

    for(i in 1:r)
    {
      reach1<-vector("character")
      anti1<-vector("character")
      inter1<-vector("character")

      reach1<-unlist(reachset[[i]])
      anti1<-unlist(antiset[[i]])
      inter1<-unlist(interset[[i]])

      reach2<-""
      anti2<-""
      inter2<-""

      for(r1 in 1:length(reach1))
      {
        reach2<-paste(reach2,reach1[r1])
      }

      for(a1 in 1:length(anti1))
      {
        anti2<-paste(anti2,anti1[a1])
      }

      for(i1 in 1:length(inter1))
      {
        inter2<-paste(inter2,inter1[i1])
      }

      final1[i,1]<-rown[i]
      final1[i,2]<-reach2
      final1[i,3]<-anti2
      final1[i,4]<-inter2
      final1[i,5]<-levelout[i]

    }

    ###


    del_vec<- vector("numeric",r)

    for (i in 1:r)
    {
      if (levelout[i]==1)
        del_vec[i]=i
    }

    C_mat_1=A_mat
    A_mat=C_mat_1[-del_vec,-del_vec]
    names(del_vec)=x
    len<- length(A_mat)

    final_mat<-rbind(final_mat,Heading,final1,b_row)


    if(len==1)
    {

      a1<-which(del_vec==0)
      x1=names(a1)
      last_row<-c(x1,x1,x1,x1,1)
      final1<-last_row
      len<-0
      final_mat<-rbind(final_mat,Heading,final1,b_row)

    }
  }


  wb1<-xlsx::createWorkbook(type = "xlsx")
  wbn<-"ISM_Output"
  sheet1 <- xlsx::createSheet(wb1, sheetName="ISM")
  xlsx::addDataFrame(x=final_mat, sheet=sheet1, characterNA="",row.names=FALSE,col.names = F)
  xlsx::autoSizeColumn(sheet1, colIndex=1:5)
  rows  <- xlsx::getRows(sheet1)
  xlsx::setRowHeight(rows, inPoints=20)
  file1 <- paste(wbn,".xlsx", sep="")
  p1<-gsub(" ","",(paste(Dir,"/",file1)))
  xlsx::saveWorkbook(wb1,p1)
  print(p1)
  options(warn=-1)
  outputformat(p1)

}

#'This Mat_format Function formats the ISM_Matrix.xlsx file That is implicitly called by ISM.
#'@param fin_mat a final matrix consists of 1s' and 0's (final reachability matrix) produced by \code{ISM}
#'@param A_mat a initial matrix consists of 1s' and 0's (initial reachability matrix) produced by \code{ISM}
#'@param file2 a final matrix consists of 1s' and 0's (final reachability matrix) produced by \code{ISM}
#'@import("xlsx")

Mat_format=function(fin_mat,A_mat,file2)
{
  r1=nrow(A_mat)
  c1=ncol(A_mat)
  wb_load<-xlsx::loadWorkbook(file2)
  sheets <- xlsx::getSheets(wb_load)
  sheet <- sheets[["New_Matrix"]]
  rows <- xlsx::getRows(sheet)
  cells <- xlsx::getCells(rows)
  count1<-c1

  for (i in 1:r1)
  {
    for(j in 1:c1)
    {
      count1<-count1+1
      if(A_mat[i,j]==0 & fin_mat[i,j]==1)
      {
        cell_select = count1
        cs <- xlsx::CellStyle(wb_load) +
          xlsx::Font(wb_load, heightInPoints=11, isBold=TRUE,isItalic=TRUE,
                     name="Courier New", color="white") +
          xlsx::Fill(backgroundColor="darkolivegreen1", foregroundColor="darkolivegreen1",
                     pattern="SOLID_FOREGROUND") +
          xlsx::Alignment(h="ALIGN_RIGHT")

        xlsx::setCellStyle(cells[[cell_select]], cs)
      }

    }

    xlsx::saveWorkbook(wb_load,file2)
    #print("Matrix Done")
  }

}

#'This outputformat Function formats the ISM_output.xlsx file that implicitly called by ISM.
#'@param file1 a Level out iterations produced by \code{ISM}
#'@import("xlsx")

outputformat=function(file1)
{
  a_1<- as.matrix(xlsx::read.xlsx(file1,sheetName = "ISM",header = FALSE))
  wb_load<-xlsx::loadWorkbook(file1)
  sheets <- xlsx::getSheets(wb_load)
  sheet <- sheets[["ISM"]]
  rows <- xlsx::getRows(sheet)
  cells <- xlsx::getCells(rows)

  for (i in 1:nrow(a_1))
  {
    a_1[[i,5]]<-as.numeric(a_1[[i,5]])

    if (is.na(a_1[[i,5]]))
    {
      paste("*")
    }
    else
    {
      if((a_1[[i,5]])==1)
      {
        a11<-paste(a_1[[i,5]])
        #print(a11)

        cell_select =(i*5)

        cs <- xlsx::CellStyle(wb_load) +
          xlsx::Font(wb_load, heightInPoints=11, isBold=TRUE,isItalic=TRUE,
                     name="Courier New", color="black") +
          xlsx::Fill(backgroundColor="forestgreen", foregroundColor="forestgreen",
                     pattern="SOLID_FOREGROUND") +
          xlsx::Alignment(h="ALIGN_RIGHT")

        xlsx::setCellStyle(cells[[cell_select]], cs)

      }
      else
      {
        paste("*")
      }
    }


  }

  xlsx::saveWorkbook(wb_load,file1)
  print("Outputs have been created")

}
