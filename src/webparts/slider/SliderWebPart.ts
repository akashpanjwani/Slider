import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IWebPartContext,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
//import { escape } from '@microsoft/sp-lodash-subset';
import * as pnp from 'sp-pnp-js';
import {
  Web
} from "sp-pnp-js";
///import styles from './Slider.module.scss';
//import * as strings from 'sliderStrings';

import { ISliderWebPartProps } from './ISliderWebPartProps';

import * as $ from 'jquery';
import * as bootstrap from 'bootstrap';
var Swiper = require('swiper');

window["$"] = $;
window["jQuery"] = $;

export interface ISlides {
  value: ISlide[];

}

export interface ISlide {
  ID: number;
  Title: string;
  Description: string;
  URL: string;
}

export interface ISlideItem {
  SPFxSliderImage: string;
  FileRef: string;
  FileLeafRef: string;
}

export interface IItemGuid {
  value: string;
}

export interface slides {
  Title: string;
  Url: string;
  Description: string;
  linkURL: any;
}
//var Swiper = require('swiper');
var count;
var desCount;
export default class SliderWebPart extends BaseClientSideWebPart<ISliderWebPartProps> {
  private _slides: slides[] = [];
  private ListName: string;

  public constructor(context: IWebPartContext) {
    super();
    SPComponentLoader.loadCss("https://desireinfowebsp.sharepoint.com/sites/Intranet/Slider/CSS/swiper.min.css");
    //SPComponentLoader.loadCss("https://chicosfas.sharepoint.com/sites/chs/SPFx_CDN/SliderDemo/css/swiper.min.css");
    SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css");
    SPComponentLoader.loadScript("https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js");
    SPComponentLoader.loadScript("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js");
  }
  public render(): void {
    this.ListName = this.properties.WebpartTitle;
  }

  public onInit<T>(): Promise<T> {
    var current = this;
    current.getSlides();
    return Promise.resolve();
  }

  protected getSlides() {

    this.ListName = this.properties.WebpartTitle ? this.properties.WebpartTitle : 'Slider Images';
    this._getSlides().then((response: ISlides): void => {

      this.domElement.innerHTML = `
        <style>
            
            .loader8{
              position: relative;
              width: 80px;
              height: 80px;
              
              top: 28%;
              top: -webkit-calc(50% - 43px);
              top: calc(50% - 43px);
              left: 35%;
              left: -webkit-calc(50% - 43px);
              left: calc(50% - 43px);
          
              border-radius: 50px;
              background-color: rgba(255, 255, 255, .2);
              border-width: 40px;
              border-style: double;
              border-color:transparent  #fff;
          
              -webkit-box-sizing:border-box;
              -moz-box-sizing:border-box;
                  box-sizing:border-box;
          
              -webkit-transform-origin:  50% 50%;
                  transform-origin:  50% 50% ;
              -webkit-animation: loader8 2s linear infinite;
                  animation: loader8 2s linear infinite;
              
          }
          @-webkit-keyframes loader8{
              0%{-webkit-transform:rotate(0deg);}
              100%{-webkit-transform:rotate(360deg);}
          }
          
          @keyframes loader8{
              0%{transform:rotate(0deg);}
              100%{transform:rotate(360deg);}
          }
          .overlay {
            height: 100%;
            width: 100%;
            position: fixed;
            z-index: 90;
            top: 0;
            left: 0;
            background-color: rgba(0,0,0, 0.9);
            overflow-y: auto;
            overflow-x: hidden;
            text-align: center;
            transition: .5s;
            display: none;
          }
          .box{
            z-index: 99;
              height: 200px;
              width: 300px;
              /*margin:0 -4px -5px -2px;*/
            transition: all .2s ease;
            margin: auto;
            margin-top: 15%;
          }
          </style>
        <div id="Loader" class="overlay">
      <div class="box">
        <div class="loader8"></div>
      </div>
      </div>
        `;
      // get slider images
      if (!this.renderedOnce) {
        this.properties.maxResultsProp.set(new Int32Array(response.value.length));
      }
      if (response.value) {

        $(".overlay").show();

        this._slides = [];
        response.value.forEach((slide: ISlide): void => {
          this._getImage(slide.ID)
            .then((data: ISlideItem): void => {
              let div = document.createElement('div');
              var fileName = data.FileLeafRef;
              const item: slides = {
                Title: slide.Title,
                Description: slide.Description,
                linkURL: slide.URL,
                Url: `${this.context.pageContext.web.absoluteUrl}/` + this.ListName + `/` + fileName
              };
              this._slides.push(item);
            })
            .then((): void => {
              this.domElement.innerHTML = `
            <style>
            #jssor_1
            {
              height:500px;
            }
            .swiper-container {
              width: 100%:
              margin-left: auto;
              margin-right: auto;
            }
            #jssor_1 a:hover,#jssor_1 a:visited,#jssor_1 a:after,#jssor_1 a:before,#jssor_1 a{
              color: #FFF !important;
            }
            .swiper-slide {
              background-size: cover;
              background-position: center;
              background-repeat: no-repeat;
            }
            .gallery-top {
              height: 80%;
              width: 100%;
            }
            .gallery-thumbs {
              height: 25%;
              box-sizing: border-box;
              padding: 10px 0;
            }
            .gallery-thumbs .swiper-slide {
              height: 100%;
              opacity: 0.4;
            }
            .gallery-thumbs .swiper-slide-active {
              opacity: 1;
            }
          .carousel-caption {
            top:0px;
            left: auto;
            right: 0px;
            max-width: 400px;
            max-height: 400px;
            overflow: hidden;
            display: block;
            background-color: #464e56;
            height: max-content; position: absolute; display: block; z-index: 1;
            padding:20px;
          }
          .propertyPaneTitleBar_b875fc01{
            height: 50px !important;
          }
          .thumbtitle
          {
            position:absolute;
            display:block;
            width:100%;
            bottom:0px;
            color:white;
            text-align:center;
            background: rgba(0, 0, 0, 0.7);
          }
          .modalDesign
          {
              font-weight: 600;
              font-size: larger;
          }
          /* Extra small devices (phones, 600px and down) */
            @media only screen and (max-width: 600px) {
              
              #jssor_1
              {
                height:200px;
              }
              .thumbtitle
              {
                font-size:10px;
              }
            } 

            /* Small devices (portrait tablets and large phones, 600px and up) */
            @media only screen and (min-width: 600px) {
              
              #jssor_1
              {
                height:300px;
              }

            }  

            /* Medium devices (landscape tablets, 768px and up) */
            @media only screen and (min-width: 768px) {
              
              #jssor_1
              {
                height:400px;
              }

            } 

            /* Large devices (laptops/desktops, 992px and up) */
            @media only screen and (min-width: 992px) {
              
              #jssor_1
              {
                height:500px;
              }

            }  

            /* Extra large devices (large laptops and desktops, 1200px and up) */
            @media only screen and (min-width: 1200px) {
              
              #jssor_1
              {
                height:400px;
              }

            } 
          </style>
          <!-- Modal -->
        <div class="modal fade" id="myModal" role="dialog">
          <div class="modal-dialog ">
          
            <!-- Modal content-->
            <div class="modal-content">
              <div class="modal-header">
                <button type="button" class="close" data-dismiss="modal">&times;</button>
              </div>
              <div class="modal-body">
                <p id="error" class="modalDesign"></p>
              </div>
            </div>
          </div>
        </div>
          <div id="jssor_1">
            <div class="swiper-container gallery-top">
              <div class="swiper-wrapper fullImg">
              </div>
              <!-- Add Arrows -->
              <div class="swiper-button-next swiper-button-white"></div>
              <div class="swiper-button-prev swiper-button-white"></div>
            </div>
            <div class="swiper-container gallery-thumbs">
              <div class="swiper-wrapper thumbImg">
              </div>
            </div>
          </div>`;

              if (this._slides.length > 0) {
                this._slides.forEach((item: slides, index: number): void => {
                  $('.thumbImg', this.domElement).append(this._itemCarouselIndicators(item, index));
                  $('.fullImg', this.domElement).append(this._itemSlideWrapper(item, index));
                });
              }

              //$('.thumbImg').find('.swiper-slide-duplicate').hide();
              var galleryTop = new Swiper('.gallery-top', {
                spaceBetween: 10,
                loop: true,
                loopedSlides: 5, //looped slides should be the same
                navigation: {
                  nextEl: '.swiper-button-next',
                  prevEl: '.swiper-button-prev',
                },
              });
              var galleryThumbs = new Swiper('.gallery-thumbs', {
                spaceBetween: 10,
                slidesPerView: 4,
                touchRatio: 0.2,
                loop: true,
                loopedSlides: 5, //looped slides should be the same
                slideToClickedSlide: true,
              });
              galleryTop.controller.control = galleryThumbs;
              galleryThumbs.controller.control = galleryTop;
            });
        });
        console.log("bind");

      }
      else {
        this.domElement.innerHTML = `
        <style>
        .modalDesign
          {
              font-weight: 600;
              font-size: larger;
          }
          </style>
        <!-- Modal -->
        <div class="modal fade" id="myModal" role="dialog">
          <div class="modal-dialog ">
          
            <!-- Modal content-->
            <div class="modal-content">
              <div class="modal-header">
                <button type="button" class="close" data-dismiss="modal">&times;</button>
              </div>
              <div class="modal-body">
                <p id="error" class="modalDesign"></p>
              </div>
            </div>
          </div>
        </div>
        `;
        $(".overlay").show();
        $('#error').text('Please create Picture library from webpart property panel');
        $('#myModal').modal("show");
      }
    });
  }

  protected onPropertyPaneConfigurationComplete(): void {
    console.log("onPropertyPaneConfigurationComplete");
  };

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private ButtonClick(oldVal: any): any {


    var current = this;
    let spWeb = new Web(this.context.pageContext.web.absoluteUrl);
    current.ListName = this.properties.WebpartTitle;
    spWeb.lists.getByTitle(current.ListName).get().then(function (result) {
      console.log('List ' + current.ListName + ' Exists');
      $('#error').text('List "' + current.ListName + '" Exists');
      $('#myModal').modal("show");
    }).catch(function (err) {


      current.domElement.innerHTML = `
      <style>
          .loader8{
            position: relative;
            width: 80px;
            height: 80px;
        
            top: 28%;
            top: -webkit-calc(50% - 43px);
            top: calc(50% - 43px);
            left: 35%;
            left: -webkit-calc(50% - 43px);
            left: calc(50% - 43px);
        
            border-radius: 50px;
            background-color: rgba(255, 255, 255, .2);
            border-width: 40px;
            border-style: double;
            border-color:transparent  #fff;
        
            -webkit-box-sizing:border-box;
            -moz-box-sizing:border-box;
                box-sizing:border-box;
        
            -webkit-transform-origin:  50% 50%;
                transform-origin:  50% 50% ;
            -webkit-animation: loader8 2s linear infinite;
                animation: loader8 2s linear infinite;
            
        }
        @-webkit-keyframes loader8{
            0%{-webkit-transform:rotate(0deg);}
            100%{-webkit-transform:rotate(360deg);}
        }
        
        @keyframes loader8{
            0%{transform:rotate(0deg);}
            100%{transform:rotate(360deg);}
        }
        .overlay {
          height: 100%;
          width: 100%;
          position: fixed;
          z-index: 90;
          top: 0;
          left: 0;
          background-color: rgba(0,0,0, 0.9);
          overflow-y: auto;
          overflow-x: hidden;
          text-align: center;
          transition: .5s;
          display: none;
        }
        .box{
          z-index: 99;
            height: 200px;
            width: 300px;
            /*margin:0 -4px -5px -2px;*/
          transition: all .2s ease;
          margin: auto;
          margin-top: 15%;
        }
        .modalDesign
          {
              font-weight: 600;
              font-size: larger;
          }
        </style>
      <div id="Loader" class="overlay">
    <div class="box">
      <div class="loader8"></div>
    </div>
    </div>
    <!-- Modal -->
    <div class="modal fade" id="documentlibraryModal" role="dialog">
      <div class="modal-dialog modal-lg">
        <!-- Modal content-->
        <div class="modal-content">
            <div class="modal-header">
              <button type="button" class="close" data-dismiss="modal">&times;</button>
            </div>
            <div class="modal-body">
              <span> Picture Library <span class="modalDesign" id="documenterror"></span> created successfully!</span></br>
              <span>Picture Library Location: </span><span id="libraryPath" class="modalDesign"></span></br>
              <p><a href="#" id="documentURL" class="modalDesign" target="_blank">Click here</a> to navigate to the picture library.<p>
            </div>
        </div>
      </div>
    </div>
      `;

      let spListTitle = current.ListName;
      $(".overlay").show();
      let spListDescription = "Slider Images";
      let spListTemplateId = 109;
      let spEnableCT = false;
      spWeb.lists.add(spListTitle, spListDescription, spListTemplateId, spEnableCT).then(function (splist) {
        console.log("New List" + spListTitle + "Created");
        current.addColumntoList();
      });
    });
  }

  private addColumntoList() {
    var current = this;
    let spWeb = new Web(current.context.pageContext.web.absoluteUrl);
    var fieldXML = "<Field DisplayName='URL' Type='URL' Required='FALSE' Name='URL' />";
    spWeb.lists.getByTitle(this.ListName).fields.createFieldAsXml(fieldXML).then(function (result) {
      console.log("column created");
      current.updateProperty();
    });
  }

  private updateProperty() {
    var arr = [];
    var current = this;
    //pnp.sp.web.getFolderByServerRelativeUrl("/sites/chs/SPFx_CDN/Slider/Images/")
    let spWeb = new Web(current.context.pageContext.web.absoluteUrl);
    //spWeb.getFolderByServerRelativeUrl("/sites/chs/SPFx_CDN/Slider/Images/")
    spWeb.getFolderByServerRelativeUrl("/sites/Intranet/SPFX1/Slider/")
      .expand("Files").get().then(r => {
        console.log(r);
        r.Files.forEach(file => {
          console.log(file.ServerRelativeUrl);
          arr.push(file);
        });
        console.log(arr);
        var filesCount = arr.length;
        var count = 0;
        current.copyDocument(arr, filesCount, count, current);
      });
  }

  private copyDocument(arr, filesCount, count, current) {
    let spWeb = new Web(current.context.pageContext.web.absoluteUrl);
    spWeb.getFileByServerRelativeUrl(arr[count].ServerRelativeUrl)
      //.copyTo('/sites/chs/' + this.ListName + '/' + arr[count].Name, true)
     .copyTo('/sites/Intranet/' + this.ListName + '/' + arr[count].Name, true)
      .then(function (res) {
        if (count < filesCount - 1) {
          count = count + 1;
          console.log(count);
          current.copyDocument(arr, filesCount, count, current)
        }
        else {
          current.UpdateField();
        }
      });
  }

  private UpdateField() {
    var current = this;
    var libraryURl = current.context.pageContext.web.absoluteUrl + "/" + current.ListName;
    var itemArr = [{ 'title': 'My New Title', 'desc': 'Here is a new description', 'Url': 'https://www.google.com' },
    { 'title': 'My New Title2', 'desc': 'Here is a new description2', 'Url': 'https://www.google.com' },
    { 'title': 'My New Title3', 'desc': 'Here is a new description3', 'Url': 'https://www.google.com' },
    { 'title': 'My New Title4', 'desc': 'Here is a new description4', 'Url': 'https://www.google.com' },
    { 'title': 'My New Title5', 'desc': 'Here is a new description5', 'Url': 'https://www.google.com' }]
    let spWeb = new Web(current.context.pageContext.web.absoluteUrl);
    spWeb.lists.getByTitle(current.ListName).items.select('FileLeafRef', 'FileRef', '*').get().then(function (result) {
      result.forEach(function (val, index) {
        spWeb.lists.getByTitle(current.ListName).items.getById(val.Id).update({
          Title: itemArr[index].title,
          Description: itemArr[index].desc,
          URL: {
            "__metadata": { type: "SP.FieldUrlValue" },
            Description: "Link",
            Url: itemArr[index].Url
          },
        }).then(r => {
          console.log(" List Created and  updated successfully!");
        });
      });
      $("#Loader").hide();
      $('#documenterror').text(current.ListName);
      $('#documentURL').attr('href', libraryURl);
      $('#libraryPath').text(libraryURl);
      $('#documentlibraryModal').modal("show");

      $('#documentlibraryModal').on('hidden.bs.modal', function () {
        current.getSlides();
      });

    });
  }
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    console.log("onPropertyPaneFieldChanged");
    switch (propertyPath) {
      case 'maxResultsProp':
        if (oldValue != newValue) {
          this._slides = [];
          this.getSlides();
        }
        break;
      case 'LengthWebpartTitle':
        if (oldValue != newValue) {
          this._slides = [];
          this.getSlides();
        }
        break;
    }
  };

  private validateDescription(value: string): string {
    if (value === null ||
      value.trim().length === 0) {
      return 'Picture library name is required';
    }
    return '';
  }
  private validateDescriptionLength(value: string): string {
    if (value === null ||
      value.trim().length === 0) {
      return 'Description length is required';
    }
    return '';
  }
  
  private validateLength(value: string): string {
    if (value === null || value.trim().length === 0) {
      return 'No of Pictures is required';
    }
    else if (value != null) {
      var pictureLen = parseInt(value);
      if (pictureLen < 4) {
        return 'No of Pictures cannot be less then 4';
      }
    }
    return '';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Create picture library and configuraition for slider"
          },
          groups: [
            {
              groupName: "Create Library",
              groupFields: [
                PropertyPaneTextField('WebpartTitle', {
                  label: "Enter picture library name",
                  value: "Slider Images",
                  onGetErrorMessage: this.validateDescription.bind(this)
                }),
                PropertyPaneButton('ClickHere',
                  {
                    text: "Create Picture Library",
                    buttonType: PropertyPaneButtonType.Primary,
                    onClick: this.ButtonClick.bind(this)
                  })
              ]
            },
            {
              groupName: "Update Slider",
              groupFields:
                [
                  PropertyPaneTextField('LengthWebpartTitle', {
                    label: "Enter Description length",
                    value: "50",
                    onGetErrorMessage: this.validateDescriptionLength.bind(this)
                  }),
                  PropertyPaneTextField('maxResultsProp', {
                    label: "Enter No of pictures",
                    value: "5",
                    onGetErrorMessage: this.validateLength.bind(this)
                  })
                ]
            }
          ]
        }
      ]
    };
  }

  private _getSlides(): Promise<ISlides> {
    this.ListName = this.properties.WebpartTitle ? this.properties.WebpartTitle : 'Slider Images';
    if (this.renderedOnce) {
      count = this.properties.maxResultsProp;

      return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('` + this.ListName + `')/items?$select=FileLeafRef,FileRef,*&$top=` + count,
        SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        });
    } else {
      return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('` + this.ListName + `')/items?$select=FileLeafRef,FileRef,*`,
        SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {

          return response.json();

        });

    }
  }

  private _getImage(id: number): Promise<ISlideItem> {

    var current = this;
    current.ListName = current.properties.WebpartTitle ? current.properties.WebpartTitle : 'Slider Images';
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('` + current.ListName + `')/items('${id}')/FieldValuesAsHtml`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getImageUrl(url: string): Promise<string> {
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${this.context.pageContext.web.serverRelativeUrl}${url}')/ListItemAllFields/ServerRelatveUrl`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((item: IItemGuid) => {
        return item.value;
      });
  }
  private add3Dots(url: string, limit: number) {
    var dots = "...";
    if (url.length > limit) {
      // you can also use substr instead of substring
      url = url.substring(0, limit) + dots;
    }
    return url;
  }

  private _itemCarouselIndicators(item: slides, index: number): string {

    var linkUrl = item.linkURL ? item.linkURL.Url : '';
    var imgTitle = item.Title ? item.Title : '';
    return `<div class="swiper-slide" style="background-image:url('${item.Url}')">
      <div class="thumbtitle"><a target="_blank" href='`+ linkUrl + `'>` + imgTitle + `</a></div>
     </div>`;

  }

  private _itemSlideWrapper(item: slides, index: number): string {
    var linkUrl = item.linkURL ? item.linkURL.Url : '';
    var imgTitle = item.Title ? item.Title : '';
    let str: string;
    desCount = parseInt(this.properties.LengthWebpartTitle);
    if (item.Description != null) {
      str = this.add3Dots(item.Description, desCount);
    }
    else {
      str = '';
    }
    return `
      <div class="swiper-slide" style="background-image:url('${item.Url}')">
      <div class="carousel-caption" style="z-index:2;">
        <span style="z-index: 1; font-size:18px; color:white;"><a target="_blank" href='`+ linkUrl + `'>` + imgTitle + `</a></span>
        <p style="z-index: 1; font-size:12px; color:white;">`+ str + `</p>
      </div>
    </div>`;

  }
}
