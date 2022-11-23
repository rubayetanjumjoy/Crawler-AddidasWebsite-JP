from selenium.webdriver.chrome.service import Service
from selenium import webdriver
  
from bs4 import BeautifulSoup as bs
import time
import pandas as pd
import openpyxl
#initialize web driver (you need to give absolute path for chormedriver.exe)
service = Service(executable_path=r"C:\Users\Rubayet Anjum Joy\.cache\selenium\chromedriver\win32\107.0.5304.62\chromedriver.exe")

def getPageInformation(url):
    try:
        driver = webdriver.Chrome(service=service)
        driver.maximize_window()
        #navigate to the url
        driver.get(f'https://shop.adidas.jp/{url}')
        #for smooth scroll
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        total_height = int(driver.execute_script("return document.body.scrollHeight"))
        for i in range(1, total_height, 5):
            driver.execute_script("window.scrollTo(0, {});".format(i)) 
        time.sleep(5)
        #Get all html elemnt
        html = driver.page_source
        #get product id 
        productID=url.replace("products/","")
        productID=productID.replace("/","")
        #beautifulsoup 
        soup = bs(html, 'html.parser')
    
      
        #get size table data
        table_body=soup.find_all('table', attrs={"class":"sizeChartTable"})
        
        data=[]
        rows = table_body[0].find_all('th')
        list=[ele.text.strip() for ele in rows]

            
        rows = table_body[1].find_all('tr')
        for row in rows:
            
            cols = row.find_all('td')
            
            cols = [ele.text.strip() for ele in cols]
            data.append([ele for ele in cols if ele]) 
        #inserting index in data array
        for i in range(len(data)):
            data[i].insert(0,list[i])
        for i in range(len(data)):
            data[i].insert(0,productID)
             
   
        df = pd.DataFrame(data) 
        

        #get breadcrrumbList
        breadcrrumbList=soup.find_all('a', attrs={"class":"breadcrumbListItemLink"  })
        breadcrumpList=[]
        
        for ele in breadcrrumbList:
            temp=ele.text
            
            if "iconArrowCircleLeft" in temp:
                x=temp.replace("iconArrowCircleLeft"," ")
                
                breadcrumpList.append(x)
            else:
                breadcrumpList.append(temp)
        #Get Catagory  
        catagory=soup.find_all('span', attrs={"class":"categoryName test-categoryName"  })
        catagoryText=""
        
        for ele in catagory:
            catagoryText=ele.text
        #get product name
        productElement=soup.find_all('h1', attrs={"class":"itemTitle test-itemTitle"  })
        productText=""
        for ele in productElement:
            productText=ele.text
        #get price
        priceElement=soup.select(".articlePrice > p:nth-child(1) > span:nth-child(1)")
        pricetext=""
    
        for ele in priceElement:
            pricetext=ele.text
        
        #get available Sizes
        sizelistdiv=soup.find_all('div', attrs={"class":"test-sizeSelector css-958jrr"})
        sizelist=[]
         
        if(sizelistdiv):
            allLi=sizelistdiv[0].find_all('button')
            
            for ele in allLi:
                if(ele['class']!=['disable']): #only get available sizes
                    sizelist.append(ele.text)
        else:
            sizelist.append("None")
        # Get image Url
        imageUrlElements=soup.select('li.slider-slide > button:nth-child(1) > div:nth-child(1) > img')
        
        imageurlList=[]
        if(imageUrlElements):
            for ele in imageUrlElements:
                imageurlList.append(ele['src'])
        else:
            imageurlList.append("None")
        
        #Get title of description
        title=soup.select('.itemFeature')
        
        #get Description 
        description=soup.select('.commentItem-mainText')
        #get special function
    
        specialfunction=soup.select('div.details:nth-child(2)')
        specialfunDes="None"
        if(specialfunction):
            specialfunDes=specialfunction[0].text
        #get product rating
    
        ratingelemnt=soup.select('#BVRRRatingOverall_ > div:nth-child(3) > span:nth-child(1)')
        rating="None"
        if(ratingelemnt):
            rating=ratingelemnt[0].text
        #Get Number of reviews
    
        reviewsElement=soup.select('.BVRRBuyAgainTotal')
        numberOfReviews="None"
        if(reviewsElement):
            numberOfReviews=reviewsElement[0].text
        #Get Recommended Rate
    
        Rateelement=soup.select('.BVRRBuyAgainPercentage > span:nth-child(1)')
    
        recommendRate="None"
        if(Rateelement):
            recommendRate=Rateelement[0].text
        
        #get keywords
        keywordlist=[]
        keywordElement=soup.select('#__next > div > div.page.css-1umyepy > div.contentsWrapper > main > div > div > div.itemTagsPosition > div > div > a')
        if(keywordElement):
            for ele in keywordElement:
                 keywordlist.append(ele.text)
        else:
            keywordlist.append({"KeyWords":"None"})
        
      
       
        df2 = pd.DataFrame([{"breadcrrumb(Catagory)":"/".join(breadcrumpList),"Catagory":catagoryText,"product name":productText,"pirce":pricetext,"Available size": " , ".join(sizelist),
        "Product Image URL":" , ".join(imageurlList), "Title of Description":title[0].text,"Description":description[0].text,"Special Function Description": specialfunDes,
        "Rating":rating,"Number Of Reviews": numberOfReviews,"Recommended Rate":recommendRate ,"Keywords":" , ".join(keywordlist),"product_ID":productID
        
        }]) 
         
        #get cordination product
        coordinateElemnts=soup.select("li.css-1gzdh76")
        coordinateProduct=[]
        if(coordinateElemnts):
            for ele in coordinateElemnts:
                dict={}
            
                dict["Coordinated Product Name"]=ele.select(" div > div > div.articleBadge.test-root.badge.test-badge.css-1p3v6yy > div ")[0].text
                dict["price"]=ele.select(" div > div.coordinate_price > div > p > span")[0].text
                dict["image link"]=ele.select("div > div > img")[0]['src']
                coordinateProduct.append(dict)
        else:
            coordinateProduct.append("None")
        coordinate= pd.DataFrame(coordinateProduct) 

        #Get Fit
    
        fitElement=soup.select('div.BVRRCustomRatingEntryWrapper:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > img')
        ratingList=[]
        fit="None"
        if(fitElement):
        
            fit=fitElement[0]['title']
            
            ratingList.append("Sense Of Fitting "+fit)
        else:
            ratingList.append("None")
        #Get Length
    
        lengthElement=soup.select('div.BVRRCustomRatingEntryWrapper:nth-child(1) > div:nth-child(2) > div:nth-child(2) > div:nth-child(2) > img:nth-child(1)')
    
        length="None"
        if(lengthElement):
        
            length=lengthElement[0]['title']
            
            ratingList.append("Appropriation of length "+length)
        else:
            ratingList.append("None")
        #Get quality of material
    
        qualityelement=soup.select('div.BVRRCustomRatingEntryWrapper:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > img:nth-child(1)')
    
        quality="None"
        if(qualityelement):
        
            quality=qualityelement[0]['title']
            
            ratingList.append("Quality of Material "+quality)
        else:
            ratingList.append("None")
        #Get Comfort
    
        comforElement=soup.select('div.BVRRCustomRatingEntryWrapper:nth-child(2) > div:nth-child(2) > div:nth-child(2) > div:nth-child(2) > img:nth-child(1)')
    
        comfort="None"
        if(comforElement):
        
            comfort=comforElement[0]['title']
            
            ratingList.append("Comfort "+comfort)
        else:
            ratingList.append("None")


        rating= pd.DataFrame({"Rating":ratingList}) 
        #comment
        
        
        commentElement=soup.select('#BVRRDisplayContentBodyID >div')
        commentList=[]
        if(commentElement):
            for ele in commentElement:
                dict={}
                ratingofuser=ele.select('#BVRRRatingOverall_Review_Display > div.BVRRRatingNormalImage > img')
                commentTitle=ele.select("div.BVRRReviewTitleContainer > span.BVRRValue.BVRRReviewTitle")
                date=ele.select( 'div:nth-child(2) > div:nth-child(1) > div:nth-child(3) > span:nth-child(2)')
                commentDescription=ele.select('.BVRRReviewTextParagraph')
                username=ele.select('#BVSubmissionPopupContainer > div.BVRRReviewDisplayStyle5BodyContent > div > div.BVRRReviewDisplayStyle5BodyUser > div.BVRRUserNicknameContainer > span.BVRRValue.BVRRUserNickname > a > span')
                dict['Username']=username[0].text
                dict['Rating Of User']=ratingofuser[0]['title']
                dict['Comment Title']=commentTitle[0].text
                dict['Comment Description']=commentDescription[0].text
                dict['Date']=date[0].text
                commentList.append(dict)
                
        else:
            commentList.append({"Comment":"None"})
        comment=pd.DataFrame(commentList)

       
        coordinate['Product_ID']=productID
        rating['Product_ID']=productID
        comment['Product_ID']=productID

        filename=url.replace("/","_")
        # with pd.ExcelWriter(f"./Products/{filename}.xlsx", engine='xlsxwriter') as writer:    
        #     df2.to_excel(writer, 'Product Page Informations')   
        #     df.to_excel(writer, 'Size')   
        #     coordinate.to_excel(writer, 'Coordinate Product')  
        #     rating.to_excel(writer, 'Rating')  
        #     comment.to_excel(writer, 'Comment')    
        
            
        #     writer.save() 
        wb = openpyxl.load_workbook('Product_Information.xlsx') # open old file
        ws = wb["Product Page Informations"] # assign sheet to work with or as below
        wr = wb["Rating"]
        wsize = wb["Size"]
        wcordinate = wb["Coordinate Product"]
        wcomment = wb["Comment"]
        
        # ws = wb.active
        #appending product page infromation sheet
        for index, row in df2.iterrows():
            print(row)
            ws.append(row.values.tolist())
        #appending rating sheet
        for index, row in rating.iterrows():
            print(row)
            wr.append(row.values.tolist())
        #appending size sheet
        for index, row in df.iterrows():
            print(row)
            wsize.append(row.values.tolist())
        #appending coordinate product sheet
        for index, row in coordinate.iterrows():
            print(row)
            wcordinate.append(row.values.tolist())
        #appending comment product sheet
        for index, row in comment.iterrows():
            print(row)
            wcomment.append(row.values.tolist())

        wb.save("Product_Information.xlsx")
        
        driver.quit()
    except:
        print(f"Error for product {url}")
        driver.quit()
 

if __name__ == "__main__":
    count=0
    df = pd.DataFrame(columns=["breadcrrumb(Catagory)","Catagory","product name","pirce","Available size","Product Image URL","Title of Description","Description","Special Function Description"
                          , "Rating","Number Of Reviews","Recommended Rate","Keywords","Product_ID"
                           ])
    newdf=df.set_index('breadcrrumb(Catagory)') 
    coordProd = pd.DataFrame(columns=["Coordinated Product Name","price","image link","Product_ID"])
    newCroodpord=coordProd.set_index("Coordinated Product Name") 

    rating = pd.DataFrame(columns=["Rating","Product_ID"])
    newRating=rating.set_index("Rating") 

    comment = pd.DataFrame(columns=["Username","Rating Of User","Comment Title","Comment Description","Date","Product_ID"])
    newCommnet=comment.set_index("Username") 

    size=pd.DataFrame(columns=["Product_ID"])
    newsize=size.set_index('Product_ID')
    
    
    writer = pd.ExcelWriter('Product_Information.xlsx', engine='xlsxwriter')
    with writer as writer:
        newdf.to_excel(writer, sheet_name='Product Page Informations')
        newsize.to_excel(writer, sheet_name='Size')
        newCroodpord.to_excel(writer, sheet_name='Coordinate Product')
        newRating.to_excel(writer, sheet_name='Rating')
        newCommnet.to_excel(writer, sheet_name='Comment')
    writer.save() 
        
    
    #looping through 1 to 8 pages for collecting over 300 products
    for i in range(1,9):
        #navigate to the url
        driver = webdriver.Chrome(service=service)
        driver.get(f'https://shop.adidas.jp/item/?gender=mens&category=wear&group=tops&page={i}')
        #for smooth scroll
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        total_height = int(driver.execute_script("return document.body.scrollHeight"))
        for i in range(1, total_height, 5):
            driver.execute_script("window.scrollTo(0, {});".format(i)) 
        time.sleep(5)
        #Get all html elemnt
        html = driver.page_source

        #beautifulsoup 
        soup = bs(html, 'html.parser')
        linkelement=soup.find_all('div', attrs={"class":"articleDisplayCard-children"})
        
        for ele in linkelement:
            temp=ele.find_all('a')
            url=temp[0]['href']
            getPageInformation(url)
            count+=1
            print(count)
    # getPageInformation('products/HB9386/')
            
      

    
        
         
  