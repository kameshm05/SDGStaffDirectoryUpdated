var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var _this = this;
import { Version } from "@microsoft/sp-core-library";
import { PropertyPaneTextField, } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as strings from "StaffdirectoryWebPartStrings";
import { SPComponentLoader } from "@microsoft/sp-loader";
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
import * as $ from "jquery";
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/photos";
import "@pnp/sp/profiles";
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.4.min.js"
// "https://code.jquery.com/jquery-3.5.1.js"
);
import "../../ExternalRef/CSS/style.css";
SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/select2@4.1.0-beta.1/dist/css/select2.min.css");
SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css");
//import "datatables";
require("datatables.net-dt");
require("datatables.net-rowgroup-dt");
SPComponentLoader.loadCss("https://cdn.datatables.net/1.10.24/css/jquery.dataTables.css");
var that;
setTimeout(function () {
    SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js");
    SPComponentLoader.loadScript("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js");
    SPComponentLoader.loadScript("https://cdn.datatables.net/1.10.24/js/jquery.dataTables.js");
    SPComponentLoader.loadCss("https://cdn.datatables.net/rowgroup/1.0.2/css/rowGroup.dataTables.min.css");
    SPComponentLoader.loadScript("https://cdn.datatables.net/rowgroup/1.0.2/js/dataTables.rowGroup.min.js");
}, 1000);
var UserDetails = [];
var listUrl = "";
var bioAttachArr = [];
var SelectedUser = "";
var ItemID = 0;
var SelectedUserProfile = [];
var selectedUsermail = "";
var CCodeHtml = "";
var CCodeArr = [];
var OfficeAddArr = [];
var AvailEditFlag = false;
var AvailEditID = 0;
var AllAvailabilityDetails = [];
var availList = [];
var StaffdirectoryWebPart = /** @class */ (function (_super) {
    __extends(StaffdirectoryWebPart, _super);
    function StaffdirectoryWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    StaffdirectoryWebPart.prototype.onInit = function () {
        var _this = this;
        return _super.prototype.onInit.call(this).then(function (_) {
            sp.setup({
                spfxContext: _this.context,
            });
            graph.setup({
                spfxContext: _this.context,
            });
        });
    };
    StaffdirectoryWebPart.prototype.render = function () {
        listUrl = this.context.pageContext.web.absoluteUrl;
        var siteindex = listUrl.toLocaleLowerCase().indexOf("sites");
        listUrl = listUrl.substr(siteindex - 1) + "/Lists/";
        this.domElement.innerHTML = "\n     <div class=\"grid-section\">    \n     <div class=\"left\">\n     <div class=\"left-nav\">\n     <div class=\"accordion\" id=\"accordionExample\">\n     <div class=\"card\">\n       <div class=\"card-header nav-items SDHEmployee show\" id=\"headingOne\">\n           <div data-toggle=\"collapse\" class=\"clsToggleCollapse\" data-target=\"#collapseOne\" aria-expanded=\"true\" aria-controls=\"collapseOne\"><span class=\"nav-icon sdh-emp\"></span>SDG Employees</div>\n           \n       </div>\n       <div id=\"collapseOne\" class=\"clsCollapse collapse\" aria-labelledby=\"headingOne\" data-parent=\"#accordionExample\">\n         <div class=\"card-body\">\n         <div class=\"filter-section\">\n         <ul>\n         <li><a href=\"#\" class=\"sdhlastnamesort\">By Last Name</a></li>\n         <li><a href=\"#\" class=\"sdhfirstnamesort\">By First Name</a></li>\n         <li><a href=\"#\" class=\"sdhLocgrouping\">By Office</a></li>\n         <li><a href=\"#\" class=\"sdhTitlgrouping\">By Title/Staff Function</a></li>\n         <li><a href=\"#\" class=\"sdhAssistantgrouping\">By Assistant</a></li>\n         </ul>\n         </div>\n         </div>\n       </div>\n     </div>\n     <div class=\"card\">\n       <div class=\"card-header nav-items  OutsidConsultant\" id=\"headingTwo\">\n           <div data-toggle=\"collapse\" class=\"clsToggleCollapse\" data-target=\"#collapseTwo\" aria-expanded=\"false\" aria-controls=\"collapseTwo\" ><span class=\"nav-icon out-con\"></span> Outside Consultant</div>\n       </div>\n       <div id=\"collapseTwo\" class=\"clsCollapse collapse\" aria-labelledby=\"headingTwo\" data-parent=\"#accordionExample\">\n         <div class=\"card-body\">\n         <div class=\"filter-section\">\n         <ul>\n         <li><a href=\"#\" class=\"OutConslastnamesort\">By Last Name</a></li>\n         <li><a href=\"#\" class=\"OutConsFirstnamesort\">By First Name</a></li>\n         <li><a href=\"#\" class=\"OutConsLocgrouping\">By Office Affiliation</a></li>\n         <li><a href=\"#\" class=\"OutConsStaffgrouping\">By Staff Function</a></li>\n         </ul>\n         </div>\n         </div>\n       </div>\n     </div>\n     <div class=\"card\">\n       <div class=\"card-header nav-items SDHAffiliates\" id=\"headingThree\">\n           <div  data-toggle=\"collapse\" class=\"clsToggleCollapse\" data-target=\"#collapseThree\" aria-expanded=\"false\" aria-controls=\"collapseThree\"><span class=\"nav-icon affli\"></span>Affiliates</div>\n       </div>\n       <div id=\"collapseThree\" class=\"clsCollapse collapse\" aria-labelledby=\"headingThree\" data-parent=\"#accordionExample\">\n         <div class=\"card-body\">\n         <div class=\"filter-section\">\n         <ul>\n         <li><a href=\"#\" class=\"Afflastnamesort\">By Last Name</a></li>\n         <li><a href=\"#\" class=\"AffFirstname\">By First Name</a></li>\n         </ul>\n         </div>\n         </div>\n       </div>\n     </div>\n     <div class=\"card\">\n       <div class=\"card-header nav-items SDHAlumini\" id=\"headingFour\">\n         <div data-toggle=\"collapse\" class=\"clsToggleCollapse\" data-target=\"#collapseFour\" aria-expanded=\"false\" aria-controls=\"collapseFour\"><span class=\"nav-icon sdh-alumini\"></span>SDG Alumni</div>\n       </div>\n       <div id=\"collapseFour\" class=\"clsCollapse collapse\" aria-labelledby=\"headingFour\" data-parent=\"#accordionExample\">\n         <div class=\"card-body\">\n         <div class=\"filter-section\">\n         <ul>\n         <li><a href=\"#\" class=\"SDHAlumniLastName\">By Last Name</a></li>\n         <li><a href=\"#\" class=\"SDHAlumniFirstName\">By First Name</a></li>\n         <li><a href=\"#\" class=\"SDHAlumniOffice\">By SDG Office</a></li>\n         </ul>\n         </div>\n         </div>\n       </div>\n     </div>\n     <div class=\"card\">\n       <div class=\"card-header nav-items SDHShowAll\" id=\"headingFive\">\n           <div data-toggle=\"collapse\" class=\"clsToggleCollapse\" data-target=\"#collapseFive\" aria-expanded=\"false\" aria-controls=\"collapseFive\"> <span class=\"nav-icon show-all\"></span>Show All People</div>\n       </div>\n       <div id=\"collapseFive\" class=\"clsCollapse collapse\" aria-labelledby=\"headingFive\" data-parent=\"#accordionExample\">\n         <div class=\"card-body\">\n         <div class=\"filter-section\">\n         <ul>\n         <li><a href=\"#\" class=\"SDHShowAllLastName\">By Last Name</a></li>\n         <li><a href=\"#\" class=\"SDHShowAllFirstName\">By First Name</a></li>\n         </ul>\n         </div>\n         </div>\n       </div>\n     </div>\n     <div class=\"card\">\n       <div class=\"card-header nav-items SDGOfficeInfo\" id=\"headingSix\">\n           <div data-toggle=\"collapse\" class=\"clsToggleCollapse\" data-target=\"#collapseSix\" aria-expanded=\"false\" aria-controls=\"collapseSix\"><span class=\"nav-icon show-office\"></span>SDG Office Info</div>\n       </div>\n       <div id=\"collapseSix\" class=\"clsCollapse collapse\" aria-labelledby=\"headingSix\" data-parent=\"#accordionExample\">\n         <div class=\"card-body\">\n         <div class=\"filter-section\">\n         <ul>\n         <li><a href=\"#\" class=\"SDGOfficeInfoLastName\">By Last Name</a></li>\n         <li><a href=\"#\"class=\"SDGOfficeInfoFirstName\">By First Name</a></li>\n         </ul>\n         </div>\n         </div>\n       </div>\n     </div>\n     <div class=\"card\">\n       <div class=\"card-header nav-items StaffAvailability\" id=\"headingSeven\">\n           <div data-toggle=\"collapse\" class=\"clsToggleCollapse\" data-target=\"#collapseSeven\" aria-expanded=\"false\" aria-controls=\"collapseSeven\"><span class=\"nav-icon staff-avail\"></span>Staff Availability</div>\n       </div>\n       <div id=\"collapseSeven\" class=\"clsCollapse collapse\" aria-labelledby=\"headingSeven\" data-parent=\"#accordionExample\">\n         <div class=\"card-body\">\n         <div class=\"filter-section\">\n         <!--<ul>\n         <li><a href=\"#\">By Office</a></li>\n         <li><a href=\"#\">By Title</a></li>\n         </ul>-->\n         </div>\n         </div>\n       </div>\n     </div> \n     <div class=\"card\">  \n       <div class=\"card-header nav-items SDGBillingRate\" id=\"headingEight\">\n           <div data-toggle=\"collapse\" class=\"clsToggleCollapse\" data-target=\"#collapseEight\" aria-expanded=\"false\" aria-controls=\"collapseEight\"><span class=\"nav-icon billing-rate\"></span>Billing Rates</div>\n       </div>\n       <div id=\"collapseEight\" class=\"clsCollapse collapse\" aria-labelledby=\"headingEight\" data-parent=\"#accordionExample\">\n         <div class=\"card-body\">\n         <div class=\"filter-section\">\n         <ul>\n         <li><a href=\"#\" class=\"SDGBillingRateLastName\">By Last Name</a></li>\n         <li><a href=\"#\" class=\"SDGBillingRateFirstName\">By First Name</a></li>\n         </ul>\n         </div>\n         </div>\n       </div>\n     </div>\n   </div> \n     </div>\n     </div>\n     <div class=\"right\">\n     <div class=\"sdh-employee\" id=\"SdhEmployeeDetails\">\n     <!-- <div class=\"title-section\">\n     <h2>Overview</h2>\n     </div> --> \n     <div class=\"title-filter-section\">\n     </div>   \n     <div class=\"sdh-emp-table oDataTable\">\n     <table  id=\"SdhEmpTable\">\n     <thead>\n     <tr>\n     <th>Name</th>\n     <th>First Name</th> \n     <th>Last Name</th>\n     <th>Phone Number</th>\n     <th>Location</th>\n     <th>Job Title</th>\n     <th>Title</th>\n     <th>Assistant</th>\n     </tr>\n     </thead>\n     <tbody id=\"SdhEmpTbody\">\n     </tbody>\n     </table>\n     </div> \n     <div class=\"sdh-outside-table oDataTable hide\">\n     <table  id=\"SdhOutsideTable\">\n     <thead>\n     <tr>\n     <th>Name</th>\n     <th>First Name</th> \n     <th>Last Name</th>\n     <th>Phone Number</th>\n     <th>Location</th>\n     <th>Job Title</th>\n     <th>Title</th>\n     <th>Assistant</th>\n     </tr>\n     </thead>\n     <tbody id=\"SdhOutsideTbody\">\n     </tbody>\n     </table>\n     </div> \n     <div class=\"sdh-Affilate-table oDataTable hide\">\n     <table  id=\"SdhAffilateTable\">\n     <thead>\n     <tr>\n     <th>Name</th>\n     <th>First Name</th> \n     <th>Last Name</th>\n     <th>Phone Number</th>\n     <th>Location</th>\n     <th>Job Title</th>\n     <th>Title</th>\n     <th>Assistant</th>\n     </tr>\n     </thead>\n     <tbody id=\"SdhAffilateTbody\">\n     </tbody>\n     </table>\n     </div>\n\n     <div class=\"sdh-Allumni-table oDataTable hide\">\n     <table  id=\"SdhAllumniTable\">\n     <thead>\n     <tr>\n     <th>Name</th>\n     <th>First Name</th> \n     <th>Last Name</th>\n     <th>Phone Number</th>\n     <th>Location</th>\n     <th>Job Title</th>\n     <th>Title</th>\n     <th>Assistant</th>\n     </tr>\n     </thead>\n     <tbody id=\"SdhAllumniTbody\">\n     </tbody>\n     </table>\n     </div>\n     <div class=\"sdh-AllPeople-table oDataTable hide\">\n     <table  id=\"SdhAllPeopleTable\">\n     <thead>\n     <tr>\n     <th>Name</th>\n     <th>First Name</th> \n     <th>Last Name</th>\n     <th>Phone Number</th>\n     <th>Location</th>\n     <th>Job Title</th>\n     <th>Title</th>\n     <th>Assistant</th>\n     </tr>\n     </thead>\n     <tbody id=\"SdhAllPeopleTbody\">\n     </tbody>\n     </table>\n     </div>\n     \n     <div class=\"sdgofficeinfotable oDataTable hide\">\n     <table  id=\"SdgofficeinfoTable\">\n     <thead>\n     <tr>\n     <th>Office</th>\n     <th>Phone</th> \n     <th>Address</th>\n     </tr>\n     </thead>\n     <tbody id=\"SdgofficeinfoTbody\">\n     </tbody>\n     </table>\n     </div>\n     <div class=\"sdgbillingrateTable oDataTable hide\">\n     <table  id=\"SdgBillingrateTable\">\n     <thead>\n     <tr>\n     <th>Name</th>\n     <th>Satff Function</th> \n     <th>Daily Rate</th>\n     <th>Hourly Rate</th>\n     <th>Effective Date</th>\n     </tr>\n     </thead>\n     <tbody id=\"SdgBillingrateTbody\">\n     </tbody>\n     </table>\n     </div>\n     <div class=\"StaffAvailabilityTable oDataTable hide\">\n     <table id=\"StaffAvailabilityTable\">\n     <thead>\n     <tr><th>User</th><th>Location</th><th>Staff Affiliates</th><th>Availability</th></tr>\n     </thead>\n     <tbody id=\"StaffAvailabilityTbody\"></tbody>\n     </table>\n     </div>\n\n     </div>\n     \n     <div class=\"user-profile-page hide\">\n     <!-- <div class=\"title-section\">\n     <h2>Employee Detail</h2>\n     </div> -->\n     <div class=\"user-profile-cover\">\n     <div class=\"cover-bg\">\n     <div class=\"profile-picture-sec\">\n     <img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAWgAAAFoCAMAAABNO5HnAAAAvVBMVEXh4eGjo6OkpKSpqamrq6vg4ODc3Nzd3d2lpaXf39/T09PU1NTBwcHOzs7ExMS8vLysrKy+vr7R0dHFxcXX19e5ubmzs7O6urrZ2dmnp6fLy8vHx8fY2NjMzMywsLDAwMDa2trV1dWysrLIyMi0tLTCwsLKysrNzc2mpqbJycnQ0NC/v7+tra2qqqrDw8OoqKjGxsa9vb3Pz8+1tbW3t7eurq7e3t62travr6+xsbHS0tK4uLi7u7vW1tbb29sZe/uLAAAG2UlEQVR4XuzcV47dSAyG0Z+KN+ccO+ecHfe/rBl4DMNtd/cNUtXD6DtLIAhCpMiSXwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIhHnfm0cVirHTam884sVu6Q1GvPkf0heq7VE+UF5bt2y97Vat+VlRniev/EVjjp12NlgdEytLWEy5G2hepDYOt7qGob2L23Dd3valPY6dsW+jvaBOKrkm2ldBVrbag+2tYeq1oX6RxYBsF6SY3vA8to8F0roRJaZmFFK2ASWA6CiT6EhuWkoQ9gablZ6l1oW47aWoF8dpvT6FrOunoD5pa7uf6CaslyV6rqD0guzYHLRK/hwJw40Cu4MUdu9Bt8C8yR4Jt+gRbmzEKvUTicFw8kY3NonOg/aJpTTf2AWWBOBTNBkvrmWF+QNDPnZoLUNOeagpKSOVdKhK550BVa5kGLOFfMCxY92ubFuYouNC9CFdyuebKrYrsyL9hcGpgnAxVaXDJPSrGKrGreVFVkU/NmykDJj1sV2Z55s0e74hwtS9k8KvNzxY8ZozvX+L67M4/uVFwT84Kt9CPz6EjFdUqgMyCjCTSHWD4cq7jOzKMzxtGu8ddwxzzaUXHFgXkTxCqwyLyJOON0j9POc/OCpbAj+hU/Zsz9Pbk2T65VbM/mybOKbd882VexjegLPXk0L154uvF/tR5N7RjJB9bvBsLEPJgI5dCcC2P5wL3QlSClJ+bYSSpIqpljh4IkpWNzapzqB3T9vCGBuGUOtWL9hDNPizMYmjND/QIloTkSJvKB4tHRK1iaE0u9hnhgDgxi/QFJZLmLEv0FvbHlbNzTG9ApWa5KHb0J9cByFNT1DhznGOngWO9CvWQ5KdX1AXweWy7Gn/Uh9CLLQdTTCkgPLLODVCshPrSMarHWgUpkGURrl2c83drWbp+0PlRebCsvFW0G+6FtLNzXxlDuXttGrrtlbQPlacvW1ppmCDPOHgJbQ/BwpmyQnh6siHVwcJoqB3iqNx/tHY/N+pPyg7Rz83Xv0n5zuff1ppPKCSS9audf1V6i9QAAAAAAAAAAAAAAAAAAAAAAEMdyAuVeZ9I4H95/uojGgf0QjKOLT/fD88ak0ysrI6SVo9qXRWgrhIsvtaNKqs2hXNlvD0LbSDho71fKWhsxvulf2NYu+jcro42d+e0isMyCxe18R2/D6HQYWY6i4elIryE9brbMgVbzONVP2G3sBeZMsNfYFf5h715302aDIADP2Lw+CIdDQhKcGuIgKKSIk1MSMND7v6zvBvqprdqY3bWfS1itRto/O+52t+KnW+2+OdSYK+5TViS9LxxqyX07p6xUeq7hXl+WPq/AX15QI+9fDryaw5d31EP7HPGqonMb5rmvYwow/upgWTDzKYQ/C2BV3o8oSNTPYVH26FEY7zGDNfnZo0DeOYclwc6jUN4ugBVxZ0HBFp0YJoxaFK41gn7ZGxWYZtDNrSOqEK0dFLscqMbhArXuIioS3UGnHw9U5uEHFCp9quOXUGfrUSFvC11cl0p1nbK+KwHs92yFYyo2DqFEsKdq+wAqhHsqtw+hQHykescY4rnvNOC7g3TPNOEZwt3QiBuINkxpRDqEZFOaMYVgTzTkCWKFGxqyCSHVkqYsIVQQ0ZQogEwJjUkgkvNpjO8g0ZzmzCHRieacIJBLaU7qIE+bBrUhz5YGbSHPmQadIc+EBk0gT48G9SDPPQ06QZ5gQ3M2AQQa0ZwRqtCExz1kClc0ZRVCqFuacguxEhqSQC53pBlHB8HyDY3Y5BDttgnoinRoQgfinZrTuxrxgeodYiiQ+1TOz6HCy4KqLV6gREHVCqjxSsVeociaaq2hyjOVeoYyXarUhTrdZs4VeaQ6j9DIdZsXEhXpU5U+1EqoSALFtlRjC9VGHlXwRlCuTKlAWkK9rEfxehkMCB8o3EMIE1yfovUdrHiKKFb0BEMuPQrVu8CU9xNFOr3DmtcFxVm8wqBsTGHGGUxya4+CeGsHqwZjijEewDAn5Rt9dOdgWzZt6kAqMm/xylpz1EI8i3hF0SxGXQxPvJrTEHXyMuVVTF9QN+WElZuUqKPiyEodC9RV+cbKvJWos0E1TbTe4wB1l89W/GSrWY4G4G4+NUHebhwEkGGYtPgpWskQAkjSXvr8x/xlGz/RKHcr/jOrXYn/1bh0Jh7/mjfpXPALjXC+O/Av7HfzEL+nERbJZME/tpgkRYg/1Mjms48Wf1PrYzbPIIBW8aDY9j/2vsef8vz9R39bDOL/2qlDIwCBGACCOMTLl4klOpP+i4MimFe7DZy7v3rcuaYqej+f3VE1K09+AgAAAAAAAAAAAAAAAAAAAAAAgBf6wsTW1jN3CAAAAABJRU5ErkJggg==\" class=\"profile-picture\"> \n     </div>  \n     <div class=\"profile-name-section\">\n     <p class=\"profile-user-name\" id=\"UserProfileName\">Sample User</p>\n     <p class=\"profile-user-mail\" id=\"UserProfileEmail\"><span class=\"user-mail-icon\"></span>Sample mail</p>        \n     </div>   \n     </div>\n     <div class=\"user-details-section\">\n     <div class=\"profile-details-left\">\n     <div class=\"user-info\">\n     <label>Job Title:</label>\n     <div class=\"title-font\" id=\"user-job-title\"></div>\n     </div>\n     <div class=\"user-info\">   \n     <label>SDG Affiliation :</label> \n     <div class=\"title-font\" id=\"user-Designation\"></div>\n     </div>\n     <div class=\"user-info\">   \n     <label>Staff Function :</label> \n     <div class=\"title-font\" id=\"user-staff-function\"></div>\n     </div>\n     \n     </div>\n     <div class=\"profile-details-right\">\n     <div class=\"user-info\"> \n     <label>Mobile:</label>\n     <div class=\"title-font\" id=\"user-phone\"></div>\n     </div>\n     \n     <div class=\"user-info hide\"><label>Personal Mail :</label><div class=\"title-font\" id=\"userpersonalmail\"></div></div> \n     <div class=\"user-info\">\n      <div class=\"d-flex align-item-center\"><label>LinkedIn :</label><div class=\"\" id=\"linkedinIDview\"></div></div>\n      </div>\n     </div> \n     </div>\n     </div>\n     <div class=\"user-profile-tabs\">\n     <div class=\"tab-section\"> \n     <div class=\"tab-header-section\">  \n     <ul class=\"nav nav-tabs\">\n       <li class=\"active\"><a data-toggle=\"tab\" href=\"#home\">Directory Information</a></li>\n       <li id=\"availabilityTab\"><a data-toggle=\"tab\" href=\"#menu1\">Availability</a></li>\n     </ul>\n     \n     </div>\n     <div>\n     <div class=\"tab-content\">\n    <div id=\"home\" class=\"tab-pane fade in active\">\n    <div class=\"text-right\" ><button class=\"btn btn-edit\" id=\"btnEdit\">Edit</button></div> \n      <div id=\"DirectoryInformation\" class=\"d-flex view-directory\">\n      <div class=\"DInfo-left col-6\">\n      <div class=\"work-address\">\n      <h4>Work Address</h4>\n      <div class=\"d-flex\"><label>Location :</label><div class=\"address-details lblRight\" id=\"WLoctionDetails\"></div></div>\n      <div class=\"d-flex align-item-center\"><label>Address:</label><div class=\"address-details lblRight\" id=\"WAddressDetails\"></div>\n      </div> \n      </div>\n      \n      <div class=\"Assistant-view\" id=\"viewAssistant\">\n      \n      </div>\n      <div class=\"personal-info\">\n      <h4>Personal Info</h4>\n      <div class=\"address-details\" id=\"PersonaInfo\"> \n      <div class=\"d-flex\"><label>Address Line:</label><div id=\"PAddLine\" class=\"lblRight\"></div></div>\n      <div class=\"d-flex\"><label>City:</label><div id=\"PAddCity\" class=\"lblRight\"></div></div>\n      <div class=\"d-flex\"><label>State:</label><div id=\"PAddState\" class=\"lblRight\"></div></div>\n      <div class=\"d-flex\"><label>Postal Code :</label><div id=\"PAddPCode\" class=\"lblRight\"></div></div> \n      <div class=\"d-flex\"><label>Country:</label><div id=\"PAddPCountry\" class=\"lblRight\"></div></div>\n      <div class=\"d-flex\"><label>Significant Other :</label><div id=\"PSignOther\" class=\"lblRight\"></div></div>\n      <div class=\"d-flex\"><label>Children :</label><div id=\"PChildren\" class=\"lblRight\"></div></div>\n      </div>\n      </div>\n      \n      \n      \n      <div class=\"StaffStatus\">\n      <h4>Staff Status</h4>\n      <p class=\"lblRight\" id=\"staffStatus\"></p> \n      <div id=\"workscheduleViewSec\">\n      <div class=\"d-flex\"><label>Work Schedule</label><p class=\"lblRight\" id=\"workSchedule\"></p></div>\n      \n      </div>\n      </div>\n      <div class=\"citizen-info\">\n      \n      <div class=\"address-details\" id=\"CitizenInfo\"> \n      <div><label>Nationality :</label><label id=\"citizenship\" class=\"lblRight\"></label></div>\n      </div>\n      </div>\n      </div>\n      <div class=\"DInfo-right col-6\">\n      <div class=\"user-billing-rates\">\n      <h4>Billing Rates</h4>\n      <div id=\"BillingRateDetails\">\n      <div class=\"billing-rates\"><label>USD Daily Rates</label><div class=\"usd-daily-rate\" id=\"UsdDailyRate\"></div></div>\n      <div class=\"billing-rates\"><label>USD Hourly Rates</label><div class=\"usd-hourly-rate\" id=\"UsdHourlyRate\"></div></div>\n      <div class=\"billing-rates\"><label>EUR Daily Rates</label><div class=\"eur-daily-rate\" id=\"EURDailyRate\"></div></div>\n      <div class=\"billing-rates\"><label>EUR Hourly Rates</label><div class=\"eur-hourly-rate\" id=\"EURHourlyRate\"></div></div>\n      <div class=\"billing-effective-date\"><label>Effective Date</label><div class=\"effective-date\" id=\"EffectiveDate\"></div></div>\n      </div>\n      </div>\n      <div class=\"Biography-Experience\"> \n      <h4>Biography and Experience</h4>\n      <div class=\"address-details\" id=\"BioExp\">  \n      <h5>Short Bio</h5>\n      <p id=\"shortbio\" class=\"lblRight\"></p> \n      <h5>Bio Attachment(s)</h5>\n      <div class=\"bio-attachment-section\" id=\"bioAttachment\"></div>\n      <div class=\"other-exp\">\n      <h5>Other Experience Details</h5> \n      <div class=\"exp\">\n      <div class=\"\"><label>Industries</label>\n      <p id=\"IndustryExp\" class=\"lblRight\"></p>\n      </div>\n      <div class=\"\"><label>Languages</label>\n      <p id=\"LanguageExp\" class=\"lblRight\"></p>\n      </div>\n      </div>\n      <div class=\"exp\">\n      <div class=\"\"><label>SDG Courses</label>\n      <p id=\"SDGCourse\" class=\"lblRight\"></p>\n      </div>\n      <div class=\"\"><label>Software</label>\n      <p id=\"SoftwareExp\" class=\"lblRight\"></p>\n      </div>\n      </div>\n      <div class=\"exp\">\n      <div class=\"\"><label>Memberships</label>\n      <p id=\"MembershipExp\" class=\"lblRight\"></p>\n      </div>\n      <div class=\"\"><label>Special Knowledge</label>\n      <p id=\"SpecialKnowledge\" class=\"lblRight\"></p>\n      </div>\n      </div>\n      </div>\n      </div>\n      </div>\n      </div>  \n      </div> \n      <div id=\"DirectoryInformationEdit\" class=\"edit-directory hide\">\n      <div class=\"d-flex\">\n      <div class=\"DInfo-left col-6\">\n      <div class=\"work-address\">\n      <h4>Work Address</h4>\n      <div class=\"address-details d-flex\" id=\"editWorAddress\">\n      <label>Location</label>\n      <div class=\"w-100\"><select id=\"workLocationDD\"></select></div>\n      </div>\n      <div class=\"Location-Addresses d-flex\">\n      <label>Location Address</label>\n      <div class=\"address-details lblRight w-100\" id=\"EditedAddressDetails\">\n\n      </div>\n      </div>\n      </div>\n      <div class=\"staff-function-edit-info\">\n      <div class=\"d-flex\">\n      <label>Staff Function</label>\n      <div class=\"w-100\"><select id=\"StaffFunctionEdit\"></select></div>\n      </div>\n      </div>\n      <div class=\"staff-affiliates-edit-info\">\n      <div class=\"d-flex\">\n      <label>Staff Affiliates</label>\n      <div class=\"w-100\"><select id=\"StaffAffiliatesEdit\"></select></div>\n      </div>\n      </div>\n      <div class=\"assisstant-info\">\n      <h4>Assisstant</h4>\n      <div class=\"assisstant-name d-flex\">\n      <label>Name</label>\n      <div class=\"w-100\"><div id=\"peoplepickerText\" title=\"APickerField\"></div></div>\n      \n      </div>\n      </div>\n      <div class=\"contact-info\">\n      <h4>Contact Info</h4>\n      <div class=\"address-details\" id=\"ContactInfo\"> \n      <div class=\"d-flex\"><label>Personal Mail :</label><div class=\"w-100\"><input type=\"text\" id=\"personalmailID\"></div></div>\n      <div class=\"d-flex\"><label>Mobile No :</label><div class=\"w-100\" id =\"mobileNoSec\"><div class=\"d-flex mobNumbers\"><select class=\"mobNoCode\"></select><input type=\"number\" class=\"mobNo\" id=\"mobileno1\"/><span class=\"addMobNo add-icon\"></span></div></div></div>\n      <div class=\"d-flex\"><label>Home No :</label><div class=\"w-100\" id=\"homeNoSec\"><div class=\"d-flex homeNumbers\"><select class=\"homeNoCode\"></select><input type=\"number\" class=\"homeno\" id=\"homeno\"/><span class=\"addHomeNo add-icon\"></span></div></div></div>\n      <div class=\"d-flex\"><label>Emergency No :</label><div class=\"w-100\" id=\"emergencyNoSec\"><div class=\"d-flex emergencyNumbers\"><select class=\"emergencyNoCode\"></select><input type=\"number\" class=\"emergencyno\" id=\"emergencyno\" /><span class=\"addEmergencyNo add-icon\"></span></div></div></div>\n      \n      <div class=\"d-flex\"><label>Significant Other :</label><div class=\"w-100\"><textarea id=\"significantOther\"></textarea></div></div>\n      <div class=\"d-flex\"><label>Children :</label><div class=\"w-100\"><textarea id=\"children\"></textarea></div></div>\n      <div class=\"d-flex\"><label>LinkedIn ID :</label><div class=\"w-100\"><input type=\"text\" id=\"linkedInID\"></div></div>\n      </div> \n      </div>\n      <div class=\"personal-info\">\n      <h4>Personal Info</h4> \n      <div class=\"address-details\" id=\"PersonaInfo\"> \n      <div class=\"d-flex\"><label>Address Line:</label><div class=\"w-100\"><input type=\"text\" id=\"PAddLineE\"></div></div>\n      <div class=\"d-flex\"><label>City:</label><div class=\"w-100\"><input type=\"text\" id=\"PAddCityE\"></div></div>\n      <div class=\"d-flex\"><label>State:</label><div class=\"w-100\"><input type=\"text\" id=\"PAddStateE\"></div></div>\n      <div class=\"d-flex\"><label>Postal Code :</label><div class=\"w-100\"><input type=\"text\" id=\"PAddPCodeE\"></div></div> \n      <div class=\"d-flex\"><label>Country:</label><div class=\"w-100\"><input type=\"text\" id=\"PAddCountryE\"></div></div>\n      </div>\n      </div>\n\n      \n      <div class=\"StaffStatus\">\n      <h4>Staff Status</h4> \n      <div class=\"d-flex w-100\"> \n      <label>Status</label><div class=\"w-100\"><select id=\"staffstatusDD\"></select></div></div>\n      <div id=\"workscheduleEdit\">\n      <div class=\"d-flex w-100 hide\" id=\"workscheduleSec\">\n      <label>Work Schedule</label>\n      <div class=\"w-100\"><input type=\"text\" id=\"workScheduleE\"></div>\n      </div>\n      </div>\n      </div> \n      <div class=\"citizen-info\">\n      <div class=\"address-details\" id=\"CitizenInfo\"> \n      <div class=\"d-flex w-100\"><label>Nationality:</label><div class=\"w-100\"><input type=\"text\" id=\"citizenshipE\"></div></div>\n      </div>\n      </div>\n      </div> \n\n      <div class=\"DInfo-right col-6\">\n      <div class=\"user-billing-rates\">\n      <h4>Billing Rates</h4>\n\n      <div id=\"BillingRateDetailsEdit\">\n      <div class=\"billing-rates\"><label>USD Daily Rates</label><div class=\"usd-daily-rate\"></div><input type=\"number\" id=\"USDDailyEdit\"/></div>\n      <div class=\"billing-rates\"><label>USD Hourly Rates</label><div class=\"usd-hourly-rate\"></div><input type=\"number\" id=\"USDHourlyEdit\" disabled/></div>\n      <div class=\"billing-rates\"><label>EUR Daily Rates</label><div class=\"eur-daily-rate\"></div><input type=\"number\" id=\"EURDailyEdit\"/></div>\n      <div class=\"billing-rates\"><label>EUR Hourly Rates</label><div class=\"eur-hourly-rate\"></div><input type=\"number\" id=\"EURHourlyEdit\" disabled/></div>\n      <div class=\"billing-rates\"><label>Other Currency</label><div class=\"eur-hourly-rate\"></div><select id=\"othercurrDD\"></select></div>\n      <div class=\"billing-rates\"><label>Daily Rate</label><div class=\"eur-hourly-rate\"></div><input type=\"number\" id=\"ODailyEdit\"/></div>\n      <div class=\"billing-rates\"><label>Hourly Rate</label><div class=\"eur-hourly-rate\"></div><input type=\"number\" id=\"OHourlyEdit\" disabled/></div>\n      <div class=\"billing-effective-date\"><label>Effective Date</label><div class=\"effective-date\"><input type=\"date\" id=\"EffectiveDateEdit\"/></div></div>\n      </div>\n      </div>\n      <div class=\"Biography-Experience\"> \n      <h4>Biography and Experience</h4>\n      <div class=\"address-details\" id=\"BioExp\">  \n      <h5>Short Bio</h5>\n      <div><textarea id=\"Eshortbio\"></textarea></div>\n      <h5>Bio Attachment(s)</h5>\n      <div class=\"bio-attachment-section\" id=\"bioAttachment\">\n      <div class=\"custom-file\">\n<input type=\"file\" name=\"myFile\" id=\"BioAttachEdit\" multiple class=\"custom-file-input\">\n<label class=\"custom-file-label\" for=\"BioAttachEdit\">Choose File</label>\n</div>\n<div class=\"quantityFilesContainer quantityFilesContainer-static\" id=\"filesfromfolder\"></div>\n<div class=\"quantityFilesContainer quantityFilesContainer-static\" id=\"otherAttachmentFiles\"></div>\n\n      </div>\n      <div class=\"other-exp\">\n      <h5>Other Experience Details</h5> \n      <div class=\"exp\">\n      <div class=\"\"><label>Industries</label>\n      <div><textarea id=\"EIndustry\"></textarea></div>\n      </div>\n      <div class=\"\"><label>Languages</label>\n      <div><textarea id=\"ELanguage\"></textarea></div>\n      </div>\n      </div>\n      <div class=\"exp\">\n      <div class=\"\"><label>SDG Courses</label>\n      <div><textarea id=\"ESDGCourse\"></textarea></div>\n      </div>\n      <div class=\"\"><label>Software</label>\n      <div><textarea id=\"ESoftwarExp\"></textarea></div>\n      </div>\n      </div>\n      <div class=\"exp\">\n      <div class=\"\"><label>Memberships</label>\n      <div><textarea id=\"EMembership\"></textarea></div>\n      </div>\n      <div class=\"\"><label>Special Knowledge</label>\n      <div><textarea id=\"ESKnowledge\"></textarea></div>\n      </div>\n      </div>\n      </div>\n      </div>\n      </div>\n      </div>\n      \n      </div>\n      <div class=\"btn-section\">\n      <button class=\"btn btn-cancel\" id=\"BtnCancel\">Cancel</button>\n      <button class=\"btn btn-submit\" id=\"BtnSubmit\">Submit</button>\n      </div>\n      </div> \n    </div>\n    <div id=\"menu1\" class=\"tab-pane fade\">\n      <div class=\"view-availability\">\n      <div class=\"availability-btn-section\">\n      <button class=\"btn btn-add-project\"  data-toggle=\"modal\" data-target=\"#addprojectmodal\">Add Project</button>\n      </div>\n      <div class=\"availability-table-section\">\n      <table id=\"UserAvailabilityTable\">\n      <thead>\n      <tr>\n      <th>Project Name</th>\n      <th>Start Date</th>\n      <th>End Date</th>\n      <th>Percentage</th>\n      <th>Comments</th>\n      <th>Action</th>\n      </tr>\n      </thead>\n      <tbody id=\"UserAvailabilityTbody\">\n      </tbody >\n      </table>\n      </div>\n      </div>\n\n      \n      <div class=\"modal fade\" id=\"addprojectmodal\" tabindex=\"-1\" role=\"dialog\" aria-labelledby=\"addprojectmodalLabel\" aria-hidden=\"true\">\n  <div class=\"modal-dialog\" role=\"document\">\n    <div class=\"modal-content\">\n      <div class=\"modal-header\">\n        <h5 class=\"modal-title\" id=\"exampleModalLabel\">Add Project</h5>\n        \n      </div>  \n      <div class=\"modal-body add-project-modal\">\n      <div class=\"d-flex\">\n      <div class=\"d-flex col-6\"><label>Project Type</label><div class=\"w-100\"><select id=\"projecttypeDD\"><option value=\"sample\">Sample</option></select></div></div>\n      <div class=\"d-flex col-6\"><label>Project Name</label><div class=\"w-100\"><input type=\"text\" id=\"projectName\" /></div></div>\n      </div>\n      <div class=\"d-flex\">\n      <div class=\"d-flex col-6\"><label>Start Date</label><div class=\"w-100\"><input type=\"date\" id=\"projectStartDate\" /></div></div>\n        <div class=\"d-flex col-6\"><label>End Date</label><div class=\"w-100\"><input type=\"date\" id=\"projectEndDate\" /></div></div>\n      </div>   \n        <div class=\"d-flex\">\n        <div class=\"d-flex col-6\"><div id=\"percentageDiv\" class=\"d-flex w-100\"><label>Percent on Project</label><div class=\"w-100\"><input type=\"text\" id=\"projectPercent\" /></div></div></div>\n        \n        </div>\n        \n        <div class=\"d-flex\">\n        <div class=\"d-flex col-6\"><label>Client</label><div class=\"w-100\"><input type=\"text\" id=\"client\" /></div></div>\n        <div class=\"d-flex col-6\"><label>Project Code</label><div class=\"w-100\"><input type=\"text\" id=\"projectCode\" /></div></div>\n        </div>\n        \n        <div class=\"d-flex\">\n        <div class=\"d-flex col-6\"><label>Practice Area</label><div class=\"w-100\"><select id=\"practiceAreaDD\"><option value=\"sample\">Sample</option></select></div></div>\n        <div class=\"d-flex col-6\"><label>Project Location</label><div class=\"w-100\"><input type=\"text\" id=\"ProjectLocation\" /></div></div>\n        </div>\n        <div class=\"d-flex\">\n        <div class=\"d-flex col-6\"><label>Availability Notes</label><div class=\"w-100\"><textarea id=\"projectAvailNotes\" ></textarea></div></div>\n        <div class=\"d-flex col-6\"><label>Comments</label><div class=\"w-100\"><textarea id=\"Projectcomments\"></textarea></div></div></div>\n      </div>\n      <div class=\"modal-footer\">\n        <button type=\"button\" class=\"btn btn-cancel\" data-dismiss=\"modal\" id=\"closeModal\">Close</button>\n        <button type=\"button\" class=\"btn btn-submit\" id=\"add-availability\">Submit</button>\n      </div>\n    </div>\n  </div>\n</div>\n\n\n    </div>\n    </div>\n    </div>\n     </div>\n     </div>\n     </div>    \n     </div>     \n     ";
        var username = document.querySelectorAll(".usernametag");
        var userpage = document.querySelector(".user-profile-page");
        var tableSection = document.querySelector(".sdh-employee");
        var viewDir = document.querySelector(".view-directory");
        var editDir = document.querySelector(".edit-directory");
        var editbtn = document.querySelector(".btn-edit");
        // ! Side Nav Click Action
        {
            $(".clsToggleCollapse").click(function () {
                $(".clsCollapse").each(function () {
                    $(this).removeClass("in").attr("style", "");
                });
                $(this).next("div").addClass("in");
            });
            onLoadData();
            ActiveSwitch();
            $(".SDHEmployee").click(function () {
                SelectedUserProfile = [];
                if (viewDir.classList["contains"]("hide")) {
                    viewDir.classList.remove("hide");
                    editDir.classList.add("hide");
                    editbtn.classList.remove("hide");
                }
                if (tableSection.classList.contains("hide")) {
                    tableSection.classList.remove("hide");
                    userpage.classList.add("hide");
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                        lengthMenu: [50, 100],
                    };
                    bindEmpTable(options);
                }
                else {
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                        lengthMenu: [50, 100],
                    };
                    bindEmpTable(options);
                }
            });
            $(".OutsidConsultant").click(function () {
                if (viewDir.classList["contains"]("hide")) {
                    viewDir.classList.remove("hide");
                    editDir.classList.add("hide");
                    editbtn.classList.remove("hide");
                }
                if (tableSection.classList.contains("hide")) {
                    tableSection.classList.remove("hide");
                    userpage.classList.add("hide");
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                        lengthMenu: [50, 100],
                    };
                    bindOutTable(options);
                }
                else {
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                        lengthMenu: [50, 100],
                    };
                    bindOutTable(options);
                }
            });
            $(".SDHAffiliates").click(function () {
                if (viewDir.classList["contains"]("hide")) {
                    viewDir.classList.remove("hide");
                    editDir.classList.add("hide");
                    editbtn.classList.remove("hide");
                }
                if (tableSection.classList.contains("hide")) {
                    tableSection.classList.remove("hide");
                    userpage.classList.add("hide");
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                        lengthMenu: [50, 100],
                    };
                    bindAffTable(options);
                }
                else {
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                        lengthMenu: [50, 100],
                    };
                    bindAffTable(options);
                }
            });
            $(".SDHAlumini").click(function () {
                if (viewDir.classList["contains"]("hide")) {
                    viewDir.classList.remove("hide");
                    editDir.classList.add("hide");
                    editbtn.classList.remove("hide");
                }
                if (tableSection.classList.contains("hide")) {
                    tableSection.classList.remove("hide");
                    userpage.classList.add("hide");
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                        lengthMenu: [50, 100],
                    };
                    bindAlumTable(options);
                }
                else {
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                        lengthMenu: [50, 100],
                    };
                    bindAlumTable(options);
                }
            });
            $(".SDHShowAll").click(function () {
                if (viewDir.classList["contains"]("hide")) {
                    viewDir.classList.remove("hide");
                    editDir.classList.add("hide");
                    editbtn.classList.remove("hide");
                }
                if (tableSection.classList.contains("hide")) {
                    tableSection.classList.remove("hide");
                    userpage.classList.add("hide");
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                    };
                    bindAllDetailTable(options);
                }
                else {
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                    };
                    bindAllDetailTable(options);
                }
            });
            $(".SDGOfficeInfo").click(function () {
                if (viewDir.classList["contains"]("hide")) {
                    viewDir.classList.remove("hide");
                    editDir.classList.add("hide");
                    editbtn.classList.remove("hide");
                }
                if (tableSection.classList.contains("hide")) {
                    tableSection.classList.remove("hide");
                    userpage.classList.add("hide");
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                    };
                    bindOfficeTable(options);
                }
                else {
                    var options = {
                        destroy: true,
                        order: [[0, "asc"]],
                    };
                    bindOfficeTable(options);
                }
            });
        }
        // Employee Filters
        $(".sdhLocgrouping").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            SdhEmpTableRowGrouping(4, "SdhEmpTable", bindEmpTable);
        });
        $(".sdhTitlgrouping").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            SdhEmpTableRowGrouping(6, "SdhEmpTable", bindEmpTable);
        });
        $(".sdhAssistantgrouping").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            SdhEmpTableRowGrouping(7, "SdhEmpTable", bindEmpTable);
        });
        $(".sdhfirstnamesort").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[1, "asc"]],
            };
            bindEmpTable(options);
        });
        $(".sdhlastnamesort").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[2, "asc"]],
            };
            bindEmpTable(options);
        });
        //OutSideConsultant
        $(".OutConslastnamesort").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[2, "asc"]],
            };
            bindOutTable(options);
        });
        $(".OutConsFirstnamesort").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[1, "asc"]],
            };
            bindOutTable(options);
        });
        $(".OutConsLocgrouping").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            SdhEmpTableRowGrouping(4, "SdhOutsideTable", bindOutTable);
        });
        $(".OutConsStaffgrouping").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            SdhEmpTableRowGrouping(6, "SdhOutsideTable", bindOutTable);
        });
        // Affliates
        $(".Afflastnamesort").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[2, "asc"]],
            };
            bindAffTable(options);
        });
        $(".AffFirstnamesort").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[1, "asc"]],
            };
            bindStaffAvailTable(options);
        });
        // Allumni
        $(".SDHAlumniLastName").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[1, "asc"]],
            };
            bindAlumTable(options);
        });
        $(".SDHAlumniFirstName").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[2, "asc"]],
            };
            bindAlumTable(options);
        });
        $(".SDHAlumniOffice").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            SdhEmpTableRowGrouping(4, "SdhAllumniTable", bindAlumTable);
        });
        // All Users
        $(".SDHShowAllLastName").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[2, "asc"]],
            };
            bindAllDetailTable(options);
        });
        $(".SDHShowAllFirstName").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[1, "asc"]],
            };
            bindAllDetailTable(options);
        });
        $(".StaffAvailability").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[0, "asc"]],
            };
            bindAllDetailTable(options);
        });
        $(".SDGBillingRate").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[0, "asc"]],
            };
            bindBillingRateTable(options);
        });
        $(".SDGBillingRateLastName").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[2, "asc"]],
            };
            bindBillingRateTable(options);
        });
        $(".SDGBillingRateFirstName").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[1, "asc"]],
            };
            bindBillingRateTable(options);
        });
        $(".SDGOfficeInfoFirstName").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[1, "asc"]],
            };
            bindOfficeTable(options);
        });
        $(".SDGOfficeInfoLastName").click(function () {
            if (viewDir.classList["contains"]("hide")) {
                viewDir.classList.remove("hide");
                editDir.classList.add("hide");
                editbtn.classList.remove("hide");
            }
            if (tableSection.classList.contains("hide")) {
                tableSection.classList.remove("hide");
                userpage.classList.add("hide");
            }
            var options = {
                destroy: true,
                order: [[2, "asc"]],
            };
            bindOfficeTable(options);
        });
        $("#btnEdit").click(function () {
            editFunction();
        });
        $("#BtnSubmit").click(function () {
            editsubmitFunction();
        });
        $("#BtnCancel").click(function () {
            editcancelFunction();
        });
        $("#add-availability").click(function () {
            if (AvailEditFlag) {
                availUpdateFunc();
            }
            else {
                availSubmitFunc();
            }
        });
        $(document).on("change", "#BioAttachEdit", function () {
            if ($(this)[0].files.length > 0) {
                for (var index = 0; index < $(this)[0].files.length; index++) {
                    var file = $("#BioAttachEdit")[0]["files"][index];
                    // if (ValidateSingleInput($("#others")[0])) {
                    bioAttachArr.push(file);
                    $("#otherAttachmentFiles").append('<div class="quantityFiles">' +
                        "<span class=upload-filename>" +
                        file.name +
                        "</span>" +
                        "<a filename='" +
                        file.name +
                        "' class=clsRemove href='#'>x</a></div>");
                    // }
                }
                $(this).val("");
                $(this).parent().find("label").text("Choose File");
            }
            // console.log(bioAttachArr);
        });
        $(document).on("click", ".clsRemove", function () {
            // console.log(bioAttachArr);
            //var filename=$(this).attr('filename');
            var filename = $(this).parent().children()[0].innerText;
            removeSelectedfile(filename);
            $(this).parent().remove();
        });
        $(document).on("click", ".remove-icon", function () {
            $(this).parent().remove();
        });
        $(document).on("click", ".action-delete", function (e) {
            var AItemID = e.currentTarget.getAttribute("data-id");
            removeAvailProject(parseInt(AItemID));
            e.currentTarget.parentElement.parentElement.parentElement.remove();
        });
        $(document).on("change", "#staffstatusDD", function () {
            if ($("#staffstatusDD").val() == "Part-time") {
                $("#workscheduleSec").removeClass("hide");
            }
            else if ($("#staffstatusDD").val() == "Full-time") {
                $("#workscheduleSec").addClass("hide");
                $("#workscheduleSec").val("");
            }
        });
        $(document).on("change", "#workLocationDD", function () {
            $("#EditedAddressDetails").html(OfficeAddArr.filter(function (add) { return $("#workLocationDD").val() == add.OfficePlace; })[0].OfficeFullAdd);
        });
        $(document).on("change", "#projecttypeDD", function () {
            if ($("#projecttypeDD").val() == "Billable/Client" || $("#projecttypeDD").val() == "Marketing") {
                $("#percentageDiv").removeClass("hide");
            }
            else {
                $("#projectPercent").val("");
                $("#percentageDiv").addClass("hide");
            }
        });
        $(document).on("click", "#editProjectAvailability", function (e) {
            AvailEditFlag = true;
            var AEditItemID = e.currentTarget.getAttribute("data-id");
            AvailEditID = AEditItemID;
            fillEditSection(AvailEditID);
        });
        $(document).on("click", "#closeModal", function () {
            AvailEditFlag = false;
            AvailEditID = 0;
            $("#projectName").val("");
            $("#projectStartDate").val("");
            $("#projectEndDate").val("");
            $("#projectPercent").val("");
            $("#practiceAreaDD").val("");
            $("#client").val("");
            $("#projectCode").val("");
            $("#ProjectLocation").val("");
            $("#projectAvailNotes").val("");
            $("#Projectcomments").val("");
        });
        $(document).on("click", ".usernametag", function () {
            if (SelectedUserProfile[0].Affiliation != "Employee" && SelectedUserProfile[0].Affiliation != "Outside Consultant") {
                $("#menu1").addClass("hide");
                $("#availabilityTab").addClass("hide");
            }
            else {
                $("#menu1").removeClass("hide");
                $("#availabilityTab").removeClass("hide");
            }
        });
    };
    Object.defineProperty(StaffdirectoryWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse("1.0");
        },
        enumerable: true,
        configurable: true
    });
    StaffdirectoryWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription,
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField("description", {
                                    label: strings.DescriptionFieldLabel,
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    };
    return StaffdirectoryWebPart;
}(BaseClientSideWebPart));
export default StaffdirectoryWebPart;
var onLoadData = function () { return __awaiter(_this, void 0, void 0, function () {
    var StaffStatusDD, LocationDD, othercurrDD, StaffFunctionDD, StaffAffiliatesDD, AvailProjectTypeDD, AvailPracticeAreaDD, LocOptionHtml, staffOptionHtml, otherCurrHtml, StaffFunHtml, StaffAffHtml, AvailProjTypeHtml, AvailPracAreaDD, listLocation, listStaffStatus, listOtherCurr, CountryCode, listStaffFunction, listStaffAff, AvailProjectType, AvailPracticeArea;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0:
                StaffStatusDD = document.querySelector("#staffstatusDD");
                LocationDD = document.querySelector("#workLocationDD");
                othercurrDD = document.querySelector("#othercurrDD");
                StaffFunctionDD = document.querySelector("#StaffFunctionEdit");
                StaffAffiliatesDD = document.querySelector("#StaffAffiliatesEdit");
                AvailProjectTypeDD = document.querySelector("#projecttypeDD");
                AvailPracticeAreaDD = document.querySelector("#practiceAreaDD");
                LocOptionHtml = "";
                staffOptionHtml = "";
                otherCurrHtml = "";
                StaffFunHtml = "";
                StaffAffHtml = "";
                AvailProjTypeHtml = "";
                AvailPracAreaDD = "";
                return [4 /*yield*/, sp.web
                        .getList(listUrl + "StaffDirectory")
                        .fields.filter("EntityPropertyName eq 'SDGOffice'")
                        .get()];
            case 1:
                listLocation = _a.sent();
                return [4 /*yield*/, sp.web
                        .getList(listUrl + "StaffDirectory")
                        .fields.filter("EntityPropertyName eq 'StaffStatus'")
                        .get()];
            case 2:
                listStaffStatus = _a.sent();
                return [4 /*yield*/, sp.web
                        .getList(listUrl + "StaffDirectory")
                        .fields.filter("EntityPropertyName eq 'OtherCurrency'")
                        .get()];
            case 3:
                listOtherCurr = _a.sent();
                return [4 /*yield*/, sp.web
                        .getList(listUrl + "StaffDirectory")
                        .fields.filter("EntityPropertyName eq 'CountryCode'")
                        .get()];
            case 4:
                CountryCode = _a.sent();
                return [4 /*yield*/, sp.web
                        .getList(listUrl + "StaffDirectory")
                        .fields.filter("EntityPropertyName eq 'stafffunction'")
                        .get()];
            case 5:
                listStaffFunction = _a.sent();
                return [4 /*yield*/, sp.web
                        .getList(listUrl + "StaffDirectory")
                        .fields.filter("EntityPropertyName eq 'SDGAffiliation'")
                        .get()];
            case 6:
                listStaffAff = _a.sent();
                return [4 /*yield*/, sp.web
                        .getList(listUrl + "SDGAvailability")
                        .fields.filter("EntityPropertyName eq 'ProjectType'")
                        .get()];
            case 7:
                AvailProjectType = _a.sent();
                return [4 /*yield*/, sp.web
                        .getList(listUrl + "SDGAvailability")
                        .fields.filter("EntityPropertyName eq 'ProjectArea'")
                        .get()];
            case 8:
                AvailPracticeArea = _a.sent();
                AvailProjectType[0]["Choices"].forEach(function (type) {
                    AvailProjTypeHtml += "<option value=\"" + type + "\">" + type + "</option>";
                });
                AvailPracticeArea[0]["Choices"].forEach(function (Area) {
                    AvailPracAreaDD += "<option value=\"" + Area + "\">" + Area + "</option>";
                });
                listLocation[0]["Choices"].forEach(function (li) {
                    LocOptionHtml += "<option value=\"" + li + "\">" + li + "</option>";
                });
                listStaffStatus[0]["Choices"].forEach(function (stff) {
                    staffOptionHtml += "<option value=\"" + stff + "\">" + stff + "</option>";
                });
                listOtherCurr[0]["Choices"].forEach(function (curr) {
                    otherCurrHtml += "<option value=\"" + curr + "\">" + curr + "</option>";
                });
                CountryCode[0]["Choices"].forEach(function (CCode) {
                    CCodeArr.push(CCode);
                    CCodeHtml += "<option value=\"" + CCode + "\">" + CCode + "</option>";
                });
                listStaffFunction[0]["Choices"].forEach(function (func) {
                    StaffFunHtml += "<option value=\"" + func + "\">" + func + "</option>";
                });
                listStaffAff[0]["Choices"].forEach(function (Aff) {
                    StaffAffHtml += "<option value=\"" + Aff + "\">" + Aff + "</option>";
                });
                AvailProjectTypeDD.innerHTML = AvailProjTypeHtml;
                AvailPracticeAreaDD.innerHTML = AvailPracAreaDD;
                LocationDD.innerHTML = LocOptionHtml;
                StaffStatusDD.innerHTML = staffOptionHtml;
                othercurrDD.innerHTML = otherCurrHtml;
                StaffFunctionDD.innerHTML = StaffFunHtml;
                StaffAffiliatesDD.innerHTML = StaffAffHtml;
                $(".mobNoCode,.homeNoCode,.emergencyNoCode").html(CCodeHtml);
                return [4 /*yield*/, sp.web
                        .getList(listUrl + "StaffDirectory")
                        .items.select("*", "User/EMail", "User/Title", "User/FirstName", "User/LastName", "User/JobTitle", "Assistant/EMail", "Assistant/Title", "User/Id")
                        .expand("User,Assistant")
                        .get()
                        .then(function (listitem) {
                        console.log(listitem);
                        listitem.forEach(function (li) {
                            UserDetails.push({
                                Name: li.User.Title != null ? li.User.Title : "Not Available",
                                FirstName: li.User.FirstName != null ? li.User.FirstName : "Not Available",
                                LastName: li.User.LastName != null ? li.User.LastName : "Not Available",
                                Usermail: li.User.EMail != null ? li.User.EMail : "Not Available",
                                UserPersonalMail: li.PersonalEmail != null ? li.PersonalEmail : "Not Available",
                                JobTitle: li.User.JobTitle != null ? li.User.JobTitle : "Not Available",
                                Assistant: li.Assistant.Title != null ? li.Assistant.Title : "Not Available",
                                AssistantMail: li.Assistant.EMail != null ? li.Assistant.EMail : "Not Available",
                                PhoneNumber: li.MobileNo != null ? li.MobileNo : "Not Available",
                                Location: li.SDGOffice != null ? li.SDGOffice : "Not Available",
                                Title: li.stafffunction != null ? li.stafffunction : "Not Available",
                                Affiliation: li.SDGAffiliation != null ? li.SDGAffiliation : "Not Available",
                                HAddLine: li.HomeAddLine != null ? li.HomeAddLine : "Not Available",
                                HAddCity: li.HomeAddCity != null ? li.HomeAddCity : "Not Available",
                                HAddState: li.HomeAddState != null ? li.HomeAddState : "Not Available",
                                HAddPCode: li.HomeAddPCode != null ? li.HomeAddPCode : "Not Available",
                                HAddPCountry: li.HomeAddCountry != null ? li.HomeAddCountry : "Not Available",
                                ShortBio: li.ShortBio != null ? li.ShortBio : "Not Available",
                                Citizen: li.Citizenship != null ? li.Citizenship : "Not Available",
                                Industry: li.IndustryExp != null ? li.IndustryExp : "Not Available",
                                Language: li.LanguageExp != null ? li.LanguageExp : "Not Available",
                                SDGCourse: li.SDGCourses != null ? li.SDGCourses : "Not Available",
                                Software: li.SoftwareExp != null ? li.SoftwareExp : "Not Available",
                                Membership: li.Membership != null ? li.Membership : "Not Available",
                                SpecialKnowledge: li.SpecialKnowledge != null ? li.SpecialKnowledge : "Not Available",
                                USDDaily: li.USDDailyRate,
                                USDHourly: li.USDHourlyRate,
                                EURDaily: li.EURDailyRate,
                                EURHourly: li.EURHourlyRate,
                                OtherCurr: li.OtherCurrency,
                                OtherCurrDaily: li.ODailyRate,
                                OtherCurrHourly: li.OHourlyRate,
                                EffectiveDate: li.EffectiveDate != null ? li.EffectiveDate : "Not Available",
                                StaffStatus: li.StaffStatus != null ? li.StaffStatus : "Not Available",
                                WorkSchedule: li.WorkingSchedule != null ? li.WorkingSchedule : "Not Available",
                                ItemID: li.ID != null ? li.ID : "Not Available",
                                LinkedInID: li.LinkedInLink != null ? li.LinkedInLink : "Not Available",
                                SignOther: li.signother != null ? li.signother : "Not Available",
                                Child: li.children != null ? li.children : "Not Available",
                                HomeNo: li.HomeNo != null ? li.HomeNo : "Not Available",
                                EmergencyNo: li.EmergencyNo != null ? li.EmergencyNo : "Not Available",
                                UserId: li.User.Id
                            });
                        });
                        getTableData();
                        console.log(UserDetails);
                    })];
            case 9:
                _a.sent();
                return [2 /*return*/];
        }
    });
}); };
var ActiveSwitch = function () {
    var navItems = document.querySelectorAll(".nav-items");
    navItems.forEach(function (li) {
        li.addEventListener("click", function (e) {
            var activeClass = document.querySelectorAll(".nav-items");
            activeClass.forEach(function (activeC) {
                activeC["classList"].remove("show");
            });
            // console.log(e.currentTarget);
            var selectedOption = e.currentTarget;
            e.currentTarget["classList"].toggle("show");
            var activeTable = document.querySelectorAll(".oDataTable");
            activeTable.forEach(function (tables) {
                if (!tables.classList.contains("hide")) {
                    tables.classList.add("hide");
                }
                selectedOption["classList"].contains("SDHEmployee")
                    ? $(".sdh-emp-table").removeClass("hide")
                    : selectedOption["classList"].contains("OutsidConsultant")
                        ? $(".sdh-outside-table").removeClass("hide")
                        : selectedOption["classList"].contains("SDHAffiliates")
                            ? $(".sdh-Affilate-table").removeClass("hide")
                            : selectedOption["classList"].contains("SDHAlumini")
                                ? $(".sdh-Allumni-table").removeClass("hide")
                                : selectedOption["classList"].contains("SDHShowAll")
                                    ? $(".sdh-AllPeople-table").removeClass("hide")
                                    : selectedOption["classList"].contains("SDGOfficeInfo")
                                        ? $(".sdgofficeinfotable").removeClass("hide")
                                        : selectedOption["classList"].contains("SDGBillingRate")
                                            ? $(".sdgbillingrateTable").removeClass("hide")
                                            : selectedOption["classList"].contains("StaffAvailability") ? $(".StaffAvailabilityTable").removeClass("hide") : "";
            });
        });
    });
};
function getTableData() {
    return __awaiter(this, void 0, void 0, function () {
        var OfficeTable, EmpTable, OutTable, AffTable, AlumTable, AllDetailsTable, BillingRateTable, OfficeDetails, AvailHtml, AvalabilityUsers, UserWithPercentage, options;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    OfficeTable = "";
                    EmpTable = "";
                    OutTable = "";
                    AffTable = "";
                    AlumTable = "";
                    AllDetailsTable = "";
                    BillingRateTable = "";
                    return [4 /*yield*/, sp.web
                            .getList(listUrl + "SDGOfficeInfo")
                            .items.get()];
                case 1:
                    OfficeDetails = _a.sent();
                    // console.log(OfficeDetails);
                    OfficeDetails.forEach(function (oDetail) {
                        OfficeTable += "<tr><td>" + oDetail.Office + "</td><td>" + (oDetail.Phone != "null" ? oDetail.Phone.split("^").join("</br>") : "") + "</td><td>" + (oDetail.Address != "null" ? oDetail.Address.split("^").join("</br>") : "") + "</td></tr>";
                    });
                    AvailHtml = "";
                    return [4 /*yield*/, sp.web.getList(listUrl + "SDGAvailability").items.select("*,UserName/Title,UserName/EMail").expand("UserName").top(5000).get()];
                case 2:
                    AllAvailabilityDetails = _a.sent();
                    AvalabilityUsers = AllAvailabilityDetails.map(function (users) {
                        return (users.UserName.EMail);
                    });
                    AvalabilityUsers = AvalabilityUsers.filter(function (item, i) { return AvalabilityUsers.indexOf(item) == i; });
                    UserWithPercentage = [];
                    AvalabilityUsers.forEach(function (users) {
                        var userPercentage = 0;
                        var UserLocation = "";
                        var UserAffiliation = "";
                        AllAvailabilityDetails.forEach(function (all) {
                            all.UserName.EMail == users ? userPercentage += parseInt(all.Percentage) : userPercentage += 0;
                        });
                        UserDetails.forEach(function (UDetails) {
                            UDetails.Usermail == users ? UserLocation = UDetails.Location : "";
                            UDetails.Usermail == users ? UserAffiliation = UDetails.Affiliation : "";
                        });
                        UserWithPercentage.push({ UserName: users, Percentage: userPercentage, Location: UserLocation, UserAff: UserAffiliation });
                        // console.log({users:users,userPercentage:userPercentage,Location:UserLocation});
                    });
                    UserWithPercentage.forEach(function (avli) {
                        AvailHtml += "<tr><td>" + avli.UserName + "</td>\n  <td>" + avli.Location + "</td>\n  <td>" + avli.UserAff + "</td>\n  <td>\n  <div class=\"d-flex align-item-center\">  \n  <div class=\"availability-progress-bar\" style=\"border: 1px solid  " + (avli.Percentage >= 50 ? "#f01616" : "#45b345") + "\">\n  <div class=\"progress-value\" style=\"height:100%;width:" + (100 - avli.Percentage) + "%; background: " + (avli.Percentage >= 50 ? "#f01616" : "#45b345") + "\"></div>\n  </div>\n  <span style=\"color:" + (avli.Percentage >= 50 ? "#f01616" : "#45b345") + "\">" + (100 - avli.Percentage) + "%</span></div>\n  </td></tr>";
                    });
                    $('#StaffAvailabilityTbody').html(AvailHtml);
                    UserDetails.forEach(function (details) {
                        // console.log(details.PhoneNumber.split("^"));
                        var ViewPhoneNumber = details.PhoneNumber.split("^");
                        ViewPhoneNumber.pop();
                        if (details.Affiliation == "Employee") {
                            EmpTable += "<tr><td class=\"user-details-td\"><div  class=\"usernametag\">" + details.Name + "</div><div class=\"HUserDetails\">\n      <img src=\"\" class=\"userimg\"/>\n      <div class=\"user-name\">" + details.Name + "</div>\n      <div class=\"user-JTitle\">" + details.JobTitle + "</div>\n      <div class=\"user-location\">" + details.Location + "</div>\n      <div class=\"user-avail-title\">Availability</div> \n      </div></td><td>" + details.FirstName + "</td><td>" + details.LastName + "</td><td>" + ViewPhoneNumber.join(",") + "</td><td>" + (details.Location == "" || details.Location == null
                                ? "Not Available"
                                : "" + details.Location) + "</td><td>" + (details.JobTitle == "" || details.JobTitle == null
                                ? "Not Available"
                                : "" + details.JobTitle) + "</td><td>" + details.Title + "</td><td>" + (details.Assistant == "" || details.Assistant == null
                                ? "Not Available"
                                : "" + details.Assistant) + "</td></tr>";
                        }
                        if (details.Affiliation == "Outside Consultant") {
                            OutTable += "<tr><td class=\"user-details-td\"><div  class=\"usernametag\">" + details.Name + "</div><div class=\"HUserDetails\">\n      <img src=\"\" class=\"userimg\"/>\n      <div class=\"user-name\">" + details.Name + "</div>\n      <div class=\"user-JTitle\">" + details.JobTitle + "</div>\n      <div class=\"user-location\">" + details.Location + "</div>\n      <div class=\"user-avail-title\">Availability</div> \n      </div></td><td>" + details.FirstName + "</td><td>" + details.LastName + "</td><td>" + ViewPhoneNumber.join(",") + "</td><td>" + (details.Location == "" || details.Location == null
                                ? "Not Available"
                                : "" + details.Location) + "</td><td>" + (details.JobTitle == "" || details.JobTitle == null
                                ? "Not Available"
                                : "" + details.JobTitle) + "</td><td>" + (details.Title == "" || details.Title == null
                                ? "Not Available"
                                : "" + details.Title) + "</td><td>" + (details.Assistant == "" || details.Assistant == null
                                ? "Not Available"
                                : "" + details.Assistant) + "</td></tr>";
                        }
                        if (details.Affiliation == "Affiliate") {
                            AffTable += "<tr><td class=\"user-details-td\"><div  class=\"usernametag\">" + details.Name + "</div><div class=\"HUserDetails\">\n      <img src=\"\" class=\"userimg\"/>\n      <div class=\"user-name\">" + details.Name + "</div>\n      <div class=\"user-JTitle\">" + details.JobTitle + "</div>\n      <div class=\"user-location\">" + details.Location + "</div>\n      <div class=\"user-avail-title\">Availability</div> \n      </div></td><td>" + details.FirstName + "</td><td>" + details.LastName + "</td><td>" + ViewPhoneNumber.join(",") + "</td><td>" + (details.Location == "" || details.Location == null
                                ? "Not Available"
                                : "" + details.Location) + "</td><td>" + (details.JobTitle == "" || details.JobTitle == null
                                ? "Not Available"
                                : "" + details.JobTitle) + "</td><td>" + (details.Title == "" || details.Title == null
                                ? "Not Available"
                                : "" + details.Title) + "</td><td>" + (details.Assistant == "" || details.Assistant == null
                                ? "Not Available"
                                : "" + details.Assistant) + "</td></tr>";
                        }
                        if (details.Affiliation == "Alumni") {
                            AlumTable += "<tr><td class=\"user-details-td\"><div  class=\"usernametag\">" + details.Name + "</div><div class=\"HUserDetails\">\n      <img src=\"\" class=\"userimg\"/>\n      <div class=\"user-name\">" + details.Name + "</div>\n      <div class=\"user-JTitle\">" + details.JobTitle + "</div>\n      <div class=\"user-location\">" + details.Location + "</div>\n      <div class=\"user-avail-title\">Availability</div> \n      </div></td><td>" + details.FirstName + "</td><td>" + details.LastName + "</td><td>" + ViewPhoneNumber.join(",") + "</td><td>" + (details.Location == "" || details.Location == null
                                ? "Not Available"
                                : "" + details.Location) + "</td><td>" + (details.JobTitle == "" || details.JobTitle == null
                                ? "Not Available"
                                : "" + details.JobTitle) + "</td><td>" + (details.Title == "" || details.Title == null
                                ? "Not Available"
                                : "" + details.Title) + "</td><td>" + (details.Assistant == "" || details.Assistant == null
                                ? "Not Available"
                                : "" + details.Assistant) + "</td></tr>";
                        }
                        AllDetailsTable += "<tr><td class=\"user-details-td\"><div  class=\"usernametag\">" + details.Name + "</div><div class=\"HUserDetails\">\n    <img src=\"\" class=\"userimg\"/>\n    <div class=\"user-name\">" + details.Name + "</div>\n    <div class=\"user-JTitle\">" + details.JobTitle + "</div>\n    <div class=\"user-location\">" + details.Location + "</div>\n    <div class=\"user-avail-title\">Availability</div> \n    </div></td><td>" + details.FirstName + "</td><td>" + details.LastName + "</td><td>" + ViewPhoneNumber.join(",") + "</td><td>" + (details.Location == "" || details.Location == null
                            ? "Not Available"
                            : "" + details.Location) + "</td><td>" + (details.JobTitle == "" || details.JobTitle == null
                            ? "Not Available"
                            : "" + details.JobTitle) + "</td><td>" + (details.Title == "" || details.Title == null
                            ? "Not Available"
                            : "" + details.Title) + "</td><td>" + (details.Assistant == "" || details.Assistant == null
                            ? "Not Available"
                            : "" + details.Assistant) + "</td></tr>";
                        BillingRateTable += "<tr><td class=\"usernametag\">" + details.Name + "</td><td>" + details.Title + "</td><td><div>" + (details.USDDaily == "" || details.USDDaily == null
                            ? ""
                            : "USD: " + details.USDDaily) + "</div><div>" + (details.EURDaily == "" || details.EURDaily == null
                            ? ""
                            : "EUR: " + details.EURDaily) + "</div><div>" + (details.OtherCurrDaily == "" || details.OtherCurrDaily == null
                            ? ""
                            : details.OtherCurr + ": " + details.OtherCurrDaily) + "</div></td><td><div>" + (details.USDHourly == "" || details.USDHourly == null
                            ? ""
                            : "USD: " + details.USDHourly) + "</div><div>" + (details.EURHourly == "" || details.EURHourly == null
                            ? ""
                            : "EUR: " + details.EURHourly) + "</div><div>" + (details.OtherCurrHourly == "" || details.OtherCurrHourly == null
                            ? ""
                            : details.OtherCurr + ": " + details.OtherCurrHourly) + "</div></td><td>" + (details.EffectiveDate == "Not Available" ? "Not Available" : new Date(details.EffectiveDate).toLocaleDateString()) + "</td></tr>";
                    });
                    $("#SdhEmpTbody").html(EmpTable);
                    $("#SdhOutsideTbody").html(OutTable);
                    $("#SdhAffilateTbody").html(AffTable);
                    $("#SdhAllumniTbody").html(AlumTable);
                    $("#SdhAllPeopleTbody").html(AllDetailsTable);
                    $("#SdgofficeinfoTbody").html(OfficeTable);
                    $("#SdgBillingrateTbody").html(BillingRateTable);
                    options = {
                        order: [[0, "asc"]],
                        lengthMenu: [50, 100],
                    };
                    bindEmpTable(options);
                    bindOutTable(options);
                    bindAffTable(options);
                    bindAlumTable(options);
                    bindAllDetailTable(options);
                    bindOfficeTable(options);
                    bindBillingRateTable(options);
                    // bindStaffAvailTable(options);
                    SdhEmpTableRowGrouping(1, "StaffAvailabilityTable", bindStaffAvailTable);
                    UserProfileDetail();
                    return [2 /*return*/];
            }
        });
    });
}
var bindEmpTable = function (options) {
    $("#SdhEmpTable").DataTable(options);
};
var bindOutTable = function (options) {
    $("#SdhOutsideTable").DataTable(options);
};
var bindAffTable = function (options) {
    $("#SdhAffilateTable").DataTable(options);
};
var bindAlumTable = function (options) {
    $("#SdhAllumniTable").DataTable(options);
};
var bindAllDetailTable = function (options) {
    $("#SdhAllPeopleTable").DataTable(options);
};
var bindOfficeTable = function (option) {
    $("#SdgofficeinfoTable").DataTable(option);
};
var bindBillingRateTable = function (option) {
    $("#SdgBillingrateTable").DataTable(option);
};
var bindStaffAvailTable = function (option) {
    $('#StaffAvailabilityTable').DataTable(option);
};
//Todo TableRowGrouping
var SdhEmpTableRowGrouping = function (colno, tablename, tablefn) {
    var collapsedGroups = {};
    var options = {
        order: [[colno, "asc"]],
        destroy: true,
        rowGroup: {
            // Uses the 'row group' plugin
            dataSrc: colno,
            startRender: function (rows, group) {
                var collapsed = !!collapsedGroups[group];
                rows.nodes().each(function (r) {
                    r.style.display = collapsed ? "none" : "";
                });
                // Add category name to the <tr>. NOTE: Hardcoded colspan
                return $("<tr/>")
                    .append('<td colspan="8">' + group + " (" + rows.count() + ")</td>")
                    .attr("data-name", group)
                    .toggleClass("collapsed", collapsed);
            },
        },
    };
    $("#" + tablename + " tbody").on("click", "tr.dtrg-start", function () {
        var name = $(this).data("name");
        collapsedGroups[name] = !collapsedGroups[name];
        // table.draw(false);
    });
    tablefn(options);
    // UserProfileDetail();
};
function startIt() {
    var schema = {};
    schema["PrincipalAccountType"] = "User,DL,SecGroup,SPGroup";
    schema["SearchPrincipalSource"] = 15;
    schema["ResolvePrincipalSource"] = 15;
    schema["AllowMultipleValues"] = false;
    schema["MaximumEntitySuggestions"] = 50;
    schema["Width"] = "280px";
    // Render and initialize the picker.
    // Pass the ID of the DOM element that contains the picker, an array of initial
    // PickerEntity objects to set the picker value, and a schema that defines
    // picker properties.
    SPClientPeoplePicker_InitStandaloneControlWrapper("peoplepickerText", null, schema);
}
var UserProfileDetail = function () { return __awaiter(_this, void 0, void 0, function () {
    var viewDir, editDir, editbtn, submitbtn, cancelbtn, office, userpage, username, sdhEmp, Edit, UserView, UserEdit, StaffStatus, WorkscheduleSection;
    var _this = this;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0:
                viewDir = document.querySelector(".view-directory");
                editDir = document.querySelector(".edit-directory");
                editbtn = document.querySelector(".edit-btn");
                submitbtn = document.querySelector("#BtnSubmit");
                cancelbtn = document.querySelector("#BtnCancel");
                ItemID = 0;
                OfficeAddArr = [];
                return [4 /*yield*/, sp.web.getList(listUrl + "SDGOfficeInfo").items.get()];
            case 1:
                office = _a.sent();
                office.forEach(function (off) {
                    OfficeAddArr.push({ OfficePlace: off.Office, OfficeFullAdd: off.Address });
                });
                SelectedUser = "";
                userpage = document.querySelector(".user-profile-page");
                username = document.querySelectorAll(".usernametag");
                sdhEmp = document.querySelector(".sdh-employee");
                Edit = document.querySelector("#btnEdit");
                UserView = document.querySelector(".view-directory");
                UserEdit = document.querySelector(".edit-directory");
                StaffStatus = document.querySelector("#staffstatusDD");
                WorkscheduleSection = document.querySelector("#workscheduleSec");
                SelectedUserProfile = [];
                username.forEach(function (btn) {
                    btn.addEventListener("click", function (e) { return __awaiter(_this, void 0, void 0, function () {
                        var specificUser, filesHtml, editfileHtml, files, billingRateHtml, html, phno, val, i, temp, finalmonth, dd, mm, yyyy, dateformat;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    if (!sdhEmp.classList.contains("hide")) {
                                        sdhEmp.classList.add("hide");
                                        userpage.classList.remove("hide");
                                    }
                                    SelectedUser = e.currentTarget["textContent"];
                                    SelectedUserProfile = UserDetails.filter(function (li) {
                                        return li.Name == SelectedUser;
                                    });
                                    useravailabilityDetails();
                                    specificUser = graph.users
                                        .getById(SelectedUserProfile[0].Usermail)
                                        .photo.getBlob()
                                        .then(function (photo) {
                                        // console.log(photo);
                                        var url = window.URL;
                                        var blobUrl = url.createObjectURL(photo);
                                        $(".profile-picture").attr("src", blobUrl);
                                    })
                                        .catch(function (err) {
                                        $(".profile-picture").attr("src", "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAWgAAAFoCAMAAABNO5HnAAAAvVBMVEXh4eGjo6OkpKSpqamrq6vg4ODc3Nzd3d2lpaXf39/T09PU1NTBwcHOzs7ExMS8vLysrKy+vr7R0dHFxcXX19e5ubmzs7O6urrZ2dmnp6fLy8vHx8fY2NjMzMywsLDAwMDa2trV1dWysrLIyMi0tLTCwsLKysrNzc2mpqbJycnQ0NC/v7+tra2qqqrDw8OoqKjGxsa9vb3Pz8+1tbW3t7eurq7e3t62travr6+xsbHS0tK4uLi7u7vW1tbb29sZe/uLAAAG2UlEQVR4XuzcV47dSAyG0Z+KN+ccO+ecHfe/rBl4DMNtd/cNUtXD6DtLIAhCpMiSXwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIhHnfm0cVirHTam884sVu6Q1GvPkf0heq7VE+UF5bt2y97Vat+VlRniev/EVjjp12NlgdEytLWEy5G2hepDYOt7qGob2L23Dd3valPY6dsW+jvaBOKrkm2ldBVrbag+2tYeq1oX6RxYBsF6SY3vA8to8F0roRJaZmFFK2ASWA6CiT6EhuWkoQ9gablZ6l1oW47aWoF8dpvT6FrOunoD5pa7uf6CaslyV6rqD0guzYHLRK/hwJw40Cu4MUdu9Bt8C8yR4Jt+gRbmzEKvUTicFw8kY3NonOg/aJpTTf2AWWBOBTNBkvrmWF+QNDPnZoLUNOeagpKSOVdKhK550BVa5kGLOFfMCxY92ubFuYouNC9CFdyuebKrYrsyL9hcGpgnAxVaXDJPSrGKrGreVFVkU/NmykDJj1sV2Z55s0e74hwtS9k8KvNzxY8ZozvX+L67M4/uVFwT84Kt9CPz6EjFdUqgMyCjCTSHWD4cq7jOzKMzxtGu8ddwxzzaUXHFgXkTxCqwyLyJOON0j9POc/OCpbAj+hU/Zsz9Pbk2T65VbM/mybOKbd882VexjegLPXk0L154uvF/tR5N7RjJB9bvBsLEPJgI5dCcC2P5wL3QlSClJ+bYSSpIqpljh4IkpWNzapzqB3T9vCGBuGUOtWL9hDNPizMYmjND/QIloTkSJvKB4tHRK1iaE0u9hnhgDgxi/QFJZLmLEv0FvbHlbNzTG9ApWa5KHb0J9cByFNT1DhznGOngWO9CvWQ5KdX1AXweWy7Gn/Uh9CLLQdTTCkgPLLODVCshPrSMarHWgUpkGURrl2c83drWbp+0PlRebCsvFW0G+6FtLNzXxlDuXttGrrtlbQPlacvW1ppmCDPOHgJbQ/BwpmyQnh6siHVwcJoqB3iqNx/tHY/N+pPyg7Rz83Xv0n5zuff1ppPKCSS9audf1V6i9QAAAAAAAAAAAAAAAAAAAAAAEMdyAuVeZ9I4H95/uojGgf0QjKOLT/fD88ak0ysrI6SVo9qXRWgrhIsvtaNKqs2hXNlvD0LbSDho71fKWhsxvulf2NYu+jcro42d+e0isMyCxe18R2/D6HQYWY6i4elIryE9brbMgVbzONVP2G3sBeZMsNfYFf5h715302aDIADP2Lw+CIdDQhKcGuIgKKSIk1MSMND7v6zvBvqprdqY3bWfS1itRto/O+52t+KnW+2+OdSYK+5TViS9LxxqyX07p6xUeq7hXl+WPq/AX15QI+9fDryaw5d31EP7HPGqonMb5rmvYwow/upgWTDzKYQ/C2BV3o8oSNTPYVH26FEY7zGDNfnZo0DeOYclwc6jUN4ugBVxZ0HBFp0YJoxaFK41gn7ZGxWYZtDNrSOqEK0dFLscqMbhArXuIioS3UGnHw9U5uEHFCp9quOXUGfrUSFvC11cl0p1nbK+KwHs92yFYyo2DqFEsKdq+wAqhHsqtw+hQHykescY4rnvNOC7g3TPNOEZwt3QiBuINkxpRDqEZFOaMYVgTzTkCWKFGxqyCSHVkqYsIVQQ0ZQogEwJjUkgkvNpjO8g0ZzmzCHRieacIJBLaU7qIE+bBrUhz5YGbSHPmQadIc+EBk0gT48G9SDPPQ06QZ5gQ3M2AQQa0ZwRqtCExz1kClc0ZRVCqFuacguxEhqSQC53pBlHB8HyDY3Y5BDttgnoinRoQgfinZrTuxrxgeodYiiQ+1TOz6HCy4KqLV6gREHVCqjxSsVeociaaq2hyjOVeoYyXarUhTrdZs4VeaQ6j9DIdZsXEhXpU5U+1EqoSALFtlRjC9VGHlXwRlCuTKlAWkK9rEfxehkMCB8o3EMIE1yfovUdrHiKKFb0BEMuPQrVu8CU9xNFOr3DmtcFxVm8wqBsTGHGGUxya4+CeGsHqwZjijEewDAn5Rt9dOdgWzZt6kAqMm/xylpz1EI8i3hF0SxGXQxPvJrTEHXyMuVVTF9QN+WElZuUqKPiyEodC9RV+cbKvJWos0E1TbTe4wB1l89W/GSrWY4G4G4+NUHebhwEkGGYtPgpWskQAkjSXvr8x/xlGz/RKHcr/jOrXYn/1bh0Jh7/mjfpXPALjXC+O/Av7HfzEL+nERbJZME/tpgkRYg/1Mjms48Wf1PrYzbPIIBW8aDY9j/2vsef8vz9R39bDOL/2qlDIwCBGACCOMTLl4klOpP+i4MimFe7DZy7v3rcuaYqej+f3VE1K09+AgAAAAAAAAAAAAAAAAAAAAAAgBf6wsTW1jN3CAAAAABJRU5ErkJggg==");
                                    });
                                    filesHtml = "";
                                    editfileHtml = "";
                                    return [4 /*yield*/, sp.web
                                            .getFolderByServerRelativeUrl("BiographyDocument/" + SelectedUserProfile[0].Usermail)
                                            .files.get()];
                                case 1:
                                    files = _a.sent();
                                    // console.log(files);
                                    files.forEach(function (file) {
                                        // console.log(file.Name.split(".").pop());
                                        if (file.Name.split(".").pop() == "doc" ||
                                            file.Name.split(".").pop() == "docx") {
                                            filesHtml += "<div class=\"doc-section\"><span class=\"word-doc\"></span><a href='" + file.ServerRelativeUrl + "' target=\"_blank\">" + file.Name + "</a></div>";
                                            editfileHtml += "<div class=\"quantityFiles\"><span class=\"upload-filename\">" + file.Name + "</span><a filename=\"" + file.Name + "\" class=\"clsfileremove\">x</a></div>";
                                        }
                                        else if (file.Name.split(".").pop() == "xlsx" || file.Name.split(".").pop() == "csv") {
                                            filesHtml += "<div class=\"doc-section\"><span class=\"excel-doc\"></span><a href='" + file.ServerRelativeUrl + "' target=\"_blank\">" + file.Name + "</a></div>";
                                            editfileHtml += "<div class=\"quantityFiles\"><span class=\"upload-filename\">" + file.Name + "</span><a  filename=\"" + file.Name + "\" class=\"clsfileremove\">x</a></div>";
                                        }
                                        else if (file.Name.split(".").pop() == "png" ||
                                            file.Name.split(".").pop() == "jpg" ||
                                            file.Name.split(".").pop() == "jpeg") {
                                            filesHtml += "<div class=\"doc-section\"><span class=\"pic-doc\"></span><a href='" + file.ServerRelativeUrl + "' target=\"_blank\">" + file.Name + "</a></div>";
                                            editfileHtml += "<div class=\"quantityFiles\"><span class=\"upload-filename\">" + file.Name + "</span><a  filename=\"" + file.Name + "\" class=\"clsfileremove\">x</a></div>";
                                        }
                                        else {
                                            filesHtml += "<div class=\"doc-section\"><span class=\"new-doc\"></span><a href='" + file.ServerRelativeUrl + "' target=\"_blank\">" + file.Name + "</a></div>";
                                            editfileHtml += "<div class=\"quantityFiles\"><span class=\"upload-filename\">" + file.Name + "</span><a  filename=\"" + file.Name + "\" class=\"clsfileremove\">x</a></div>";
                                        }
                                    });
                                    billingRateHtml = "";
                                    if (SelectedUserProfile[0].USDDaily != null && SelectedUserProfile[0].USDDaily != 0 && SelectedUserProfile[0].USDDaily != "0") {
                                        billingRateHtml += "<div class=\"billing-rates\"><label>USD Daily Rates</label><div class=\"usd-daily-rate lblBlue\" id=\"UsdDailyRate\">" + SelectedUserProfile[0].USDDaily + "</div></div><div class=\"billing-rates\"><label>USD Hourly Rates</label><div class=\"usd-hourly-rate lblBlue\" id=\"UsdHourlyRate\">" + SelectedUserProfile[0].USDHourly + "</div></div>";
                                    }
                                    if (SelectedUserProfile[0].EURDaily != null &&
                                        SelectedUserProfile[0].EURDaily != 0 &&
                                        SelectedUserProfile[0].EURDaily != "0") {
                                        billingRateHtml += "<div class=\"billing-rates\"><label>EUR Daily Rates</label><div class=\"eur-daily-rate lblBlue\" id=\"EURDailyRate\">" + SelectedUserProfile[0].EURDaily + "</div></div><div class=\"billing-rates\"><label>EUR Hourly Rates</label><div class=\"eur-hourly-rate lblBlue\" id=\"EURHourlyRate\">" + SelectedUserProfile[0].EURHourly + "</div></div>";
                                    }
                                    if (SelectedUserProfile[0].OtherCurrDaily != null &&
                                        SelectedUserProfile[0].OtherCurrDaily != 0 &&
                                        SelectedUserProfile[0].OtherCurrDaily != "0") {
                                        billingRateHtml += "<div class=\"billing-rates\"><label>" + SelectedUserProfile[0].OtherCurr + " Daily Rates</label><div class=\"eur-daily-rate lblBlue\" id=\"oDailyRate\">" + SelectedUserProfile[0].OtherCurrDaily + "</div></div><div class=\"billing-rates\"><label>" + SelectedUserProfile[0].OtherCurr + " Hourly Rates</label><div class=\"eur-hourly-rate lblBlue\" id=\"oHourlyRate\">" + SelectedUserProfile[0].OtherCurrHourly + "</div></div>";
                                    }
                                    if (SelectedUserProfile[0].EffectiveDate != null) {
                                        billingRateHtml += " <div class=\"billing-effective-date\"><label>Effective Date</label><div class=\"effective-date lblBlue\" id=\"EffectiveDate\">" + new Date(SelectedUserProfile[0].EffectiveDate).toLocaleDateString() + "</div></div>";
                                    }
                                    ItemID = SelectedUserProfile[0].ItemID;
                                    selectedUsermail = SelectedUserProfile[0].Usermail;
                                    // console.log(selectedUsermail);
                                    if (SelectedUserProfile[0].UserPersonalMail != "" && SelectedUserProfile[0].UserPersonalMail != null) {
                                        $('#userpersonalmail').parent().removeClass('hide');
                                        $('#userpersonalmail').html(SelectedUserProfile[0].UserPersonalMail);
                                    }
                                    else {
                                        $('#userpersonalmail').html("");
                                        $('#userpersonalmail').parent().addClass('hide');
                                    }
                                    if (SelectedUserProfile[0].Assistant != null && SelectedUserProfile[0].Assistant != "") {
                                        $("#viewAssistant").html("<div class=\"d-flex align-item-center\">\n        <label>Assistant : </label><div class=\"lblRight\" id=\"assistantViewpage\">" + SelectedUserProfile[0].Assistant + "</div>\n        </div>");
                                    }
                                    else {
                                        $("#viewAssistant").html("");
                                    }
                                    if (SelectedUserProfile[0].PhoneNumber) {
                                        html = "";
                                        phno = SelectedUserProfile[0].PhoneNumber;
                                        val = phno.split("^");
                                        console.log("Split");
                                        console.log(val);
                                        if (val.length > 1) {
                                            for (i = 0; i < val.length - 1; i++) {
                                                temp = val[i].split("-");
                                                if (temp[1] == " ") {
                                                    html += "Not Available";
                                                }
                                                else {
                                                    html += val[i] + ";";
                                                }
                                            }
                                            $("#user-phone").html(html);
                                        }
                                        else {
                                            $("#user-phone").val("Not Available");
                                        }
                                    }
                                    // $('#linkedinIDview').html(`<span class="linkedInBtn" id="linkedInBtn">Link</span>`);
                                    $('#linkedinIDview').html("<a href=\"" + SelectedUserProfile[0].LinkedInID.Url + "\" target ='_blank' data-interception=\"off\"><span class=\"icon-linkedin\"></span></a>");
                                    $('#PSignOther').html(SelectedUserProfile[0].SignOther);
                                    $('#PChildren').html(SelectedUserProfile[0].Child);
                                    $("#user-Designation").html(SelectedUserProfile[0].Affiliation);
                                    $("#user-staff-function").html(SelectedUserProfile[0].Title);
                                    $("#user-job-title").html(SelectedUserProfile[0].JobTitle);
                                    $("#user-location").html(SelectedUserProfile[0].Location);
                                    // $("#user-office").html(SelectedUserProfile[0].Location);
                                    // $("#user-phone").html(SelectedUserProfile[0].PhoneNumber);
                                    // $("#user-mail").html(SelectedUserProfile[0].Usermail);
                                    // $("#personal-mail").html(SelectedUserProfile[0].UserPersonalMail);
                                    $("#UserProfileName").html(SelectedUserProfile[0].Name);
                                    $("#UserProfileEmail").html("<span class=\"user-mail-icon\"></span>" + SelectedUserProfile[0].Usermail);
                                    $("#PAddLine").html(SelectedUserProfile[0].HAddLine);
                                    $("#PAddCity").html(SelectedUserProfile[0].HAddCity);
                                    $("#PAddState").html(SelectedUserProfile[0].HAddState);
                                    $("#PAddPCode").html(SelectedUserProfile[0].HAddPCode);
                                    $("#PAddPCountry").html(SelectedUserProfile[0].HAddPCountry);
                                    $("#WAddressDetails").html(OfficeAddArr.filter(function (add) { return SelectedUserProfile[0].Location == add.OfficePlace; })[0].OfficeFullAdd);
                                    $("#WLoctionDetails").html(SelectedUserProfile[0].Location);
                                    $("#shortbio").html(SelectedUserProfile[0].ShortBio);
                                    $("#citizenship").html(SelectedUserProfile[0].Citizen);
                                    $("#IndustryExp").html(SelectedUserProfile[0].Industry);
                                    $("#LanguageExp").html(SelectedUserProfile[0].Language);
                                    $("#SDGCourse").html(SelectedUserProfile[0].SDGCourse);
                                    $("#SoftwareExp").html(SelectedUserProfile[0].Software);
                                    $("#MembershipExp").html(SelectedUserProfile[0].Membership);
                                    $("#SpecialKnowledge").html(SelectedUserProfile[0].SpecialKnowledge);
                                    $("#BillingRateDetails").html(billingRateHtml);
                                    $("#bioAttachment").html(filesHtml);
                                    $("#filesfromfolder").html(editfileHtml);
                                    $("#staffStatus").html(SelectedUserProfile[0].StaffStatus);
                                    $("#workscheduleViewSec").html(SelectedUserProfile[0].StaffStatus == "Part-time"
                                        ? "\n      <div class=\"d-flex\"><label>Work Schedule</label><p class=\"lblRight\" id=\"workSchedule\">" + (SelectedUserProfile[0].WorkSchedule == null || SelectedUserProfile[0].WorkSchedule == "" ? "Not Available" : SelectedUserProfile[0].WorkSchedule) + "</p></div>"
                                        : "");
                                    finalmonth = "";
                                    dd = new Date(SelectedUserProfile[0].EffectiveDate).getDate();
                                    mm = new Date(SelectedUserProfile[0].EffectiveDate).getMonth() + 1;
                                    mm < 10 ? (finalmonth = "0" + mm) : (finalmonth = mm);
                                    yyyy = new Date(SelectedUserProfile[0].EffectiveDate).getFullYear();
                                    dateformat = yyyy + "-" + finalmonth + "-" + dd;
                                    $("#EffectiveDateEdit").val(dateformat);
                                    return [2 /*return*/];
                            }
                        });
                    }); });
                });
                $("#USDDailyEdit").keyup(function () {
                    var usdvalue = $("#USDDailyEdit").val();
                    var finalusdval = usdvalue / 8;
                    $("#USDHourlyEdit").val(finalusdval);
                });
                $("#EURDailyEdit").keyup(function () {
                    var eurdaily = $("#EURDailyEdit").val();
                    var finaleurval = eurdaily / 8;
                    $("#EURHourlyEdit").val(finaleurval);
                });
                $("#ODailyEdit").keyup(function () {
                    var ovalue = $("#ODailyEdit").val();
                    var finalovalue = ovalue / 8;
                    $("#OHourlyEdit").val(finalovalue);
                });
                $(document).on("click", ".clsfileremove", function () {
                    var filename = $(this).attr("filename");
                    $(this).parent().remove();
                    sp.web
                        .getFileByServerRelativeUrl("/sites/StaffDirectory/BiographyDocument/" + SelectedUserProfile[0].Usermail + "/" + filename)
                        .recycle()
                        .then(function (data) { });
                });
                return [2 /*return*/];
        }
    });
}); };
var editFunction = function () { return __awaiter(_this, void 0, void 0, function () {
    var Edit, UserView, UserEdit, MobileNumberHtmlSec, HomeNumberHtmlSec, EmergencyNumberHtmlSec, MCCodeArr, AllMnumber, AllMobileNumbers, HCCodeArr, AllHNumber, AllHomeNumber, ECCodeArr, AllENumber, AllEmergencyNumber, emailAddress, divID;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0: return [4 /*yield*/, SPComponentLoader.loadScript("/_layouts/15/init.js").then(function () { })];
            case 1:
                _a.sent();
                return [4 /*yield*/, SPComponentLoader.loadScript("/_layouts/15/1033/sts_strings.js")];
            case 2:
                _a.sent();
                return [4 /*yield*/, SPComponentLoader.loadScript("/_layouts/15/clientforms.js")];
            case 3:
                _a.sent();
                return [4 /*yield*/, SPComponentLoader.loadScript("/_layouts/15/clienttemplates.js")];
            case 4:
                _a.sent();
                return [4 /*yield*/, SPComponentLoader.loadScript("/_layouts/15/clientpeoplepicker.js")];
            case 5:
                _a.sent();
                return [4 /*yield*/, SPComponentLoader.loadScript("/_layouts/15/autofill.js")];
            case 6:
                _a.sent();
                return [4 /*yield*/, SPComponentLoader.loadScript("/_layouts/15/SP.js")];
            case 7:
                _a.sent();
                return [4 /*yield*/, SPComponentLoader.loadScript("/_layouts/15/sp.runtime.js")];
            case 8:
                _a.sent();
                return [4 /*yield*/, SPComponentLoader.loadScript("/_layouts/15/sp.core.js")];
            case 9:
                _a.sent();
                return [4 /*yield*/, startIt()];
            case 10:
                _a.sent();
                Edit = document.querySelector("#btnEdit");
                UserView = document.querySelector(".view-directory");
                UserEdit = document.querySelector(".edit-directory");
                if (!UserView.classList.contains("hide")) {
                    UserView.classList.add("hide");
                    UserEdit.classList.remove("hide");
                    Edit.classList.add("hide");
                }
                else {
                    UserEdit.classList.remove("hide");
                    Edit.classList.add("hide");
                }
                MobileNumberHtmlSec = "";
                HomeNumberHtmlSec = "";
                EmergencyNumberHtmlSec = "";
                MCCodeArr = [];
                if (SelectedUserProfile[0].PhoneNumber != "" && SelectedUserProfile[0].PhoneNumber != null) {
                    AllMnumber = SelectedUserProfile[0].PhoneNumber.split("^");
                    AllMnumber.pop();
                    AllMobileNumbers = AllMnumber;
                    AllMobileNumbers.forEach(function (numbers, i) {
                        // console.log(CCodeHtml)
                        var SplitedMNum = numbers.split(" - ");
                        MCCodeArr.push(SplitedMNum[0]);
                        if (i == 0) {
                            MobileNumberHtmlSec += "<div class=\"d-flex mobNumbers\"><select class=\"mobNoCode\">" + CCodeHtml + "</select><input type=\"number\" class=\"mobNo\" id=\"\" value=\"" + SplitedMNum[1] + "\"><span class=\"addMobNo add-icon\"></div>";
                        }
                        else {
                            MobileNumberHtmlSec += "<div class=\"d-flex mobNumbers\"><select class=\"mobNoCode\">" + CCodeHtml + "</select><input type=\"number\" class=\"mobNo\" id=\"\" value=\"" + SplitedMNum[1] + "\"><span class=\"removeMobNo remove-icon\"></div>";
                        }
                        $("#mobileNoSec").html(MobileNumberHtmlSec);
                    });
                }
                else {
                    MobileNumberHtmlSec += "<div class=\"d-flex mobNumbers\"><select class=\"mobNoCode\">" + CCodeHtml + "</select><input type=\"number\" class=\"mobNo\" id=\"\"><span class=\"addMobNo add-icon\"></div>";
                    $("#mobileNoSec").html(MobileNumberHtmlSec);
                }
                HCCodeArr = [];
                if (SelectedUserProfile[0].HomeNo != "" && SelectedUserProfile[0].HomeNo != null) {
                    AllHNumber = SelectedUserProfile[0].HomeNo.split("^");
                    AllHNumber.pop();
                    AllHomeNumber = AllHNumber;
                    AllHomeNumber.forEach(function (hnumbs, j) {
                        var SplitedHNum = hnumbs.split(' - ');
                        HCCodeArr.push(SplitedHNum[0]);
                        if (j == 0) {
                            HomeNumberHtmlSec += "<div class=\"d-flex homeNumbers\"><select class=\"homeNoCode\">" + CCodeHtml + "</select><input type=\"number\" class=\"home\" id=\"\" value=\"" + SplitedHNum[1] + "\"><span class=\"addHomeNo add-icon\"></div>";
                        }
                        else {
                            HomeNumberHtmlSec += "<div class=\"d-flex homeNumbers\"><select class=\"homeNoCode\">" + CCodeHtml + "</select><input type=\"number\" class=\"home\" id=\"\" value=\"" + SplitedHNum[1] + "\"><span class=\"removeHomeNo remove-icon\"></div>";
                        }
                    });
                    $('#homeNoSec').html(HomeNumberHtmlSec);
                }
                else {
                    HomeNumberHtmlSec += "<div class=\"d-flex homeNumbers\"><select class=\"homeNoCode\">" + CCodeHtml + "</select><input type=\"number\" class=\"home\" id=\"\"><span class=\"addHomeNo add-icon\"></div>";
                    $('#homeNoSec').html(HomeNumberHtmlSec);
                }
                ECCodeArr = [];
                if (SelectedUserProfile[0].EmergencyNo != "" && SelectedUserProfile[0].EmergencyNo != null) {
                    AllENumber = SelectedUserProfile[0].EmergencyNo.split("^");
                    AllENumber.pop();
                    AllEmergencyNumber = AllENumber;
                    AllEmergencyNumber.forEach(function (enums, k) {
                        var SplitedENum = enums.split(' - ');
                        ECCodeArr.push(SplitedENum[0]);
                        if (k == 0) {
                            EmergencyNumberHtmlSec += "<div class=\"d-flex emergencyNumbers\"><select class=\"emergencyNoCode\" id=\"ec" + k + "\">" + CCodeHtml + "</select><input type=\"number\" class=\"home\" id=\"\" value=\"" + SplitedENum[1] + "\"><span class=\"addEmergencyNo add-icon\"></div>";
                        }
                        else {
                            EmergencyNumberHtmlSec += "<div class=\"d-flex emergencyNumbers\"><select class=\"emergencyNoCode\" id=\"ec" + k + "\">" + CCodeHtml + "</select><input type=\"number\" class=\"home\" id=\"\" value=\"" + SplitedENum[1] + "\"><span class=\"removeEmergencyNo remove-icon\"></div>";
                        }
                    });
                    $('#emergencyNoSec').html(EmergencyNumberHtmlSec);
                }
                else {
                    EmergencyNumberHtmlSec += "<div class=\"d-flex emergencyNumbers\"><select class=\"emergencyNoCode\">" + CCodeHtml + "</select><input type=\"number\" class=\"home\" id=\"\"><span class=\"addEmergencyNo add-icon\"></div>";
                    $('#emergencyNoSec').html(EmergencyNumberHtmlSec);
                }
                if (ECCodeArr.length > 0) {
                    $('.emergencyNoCode').each(function (i, evt) {
                        var ecID = ECCodeArr[i];
                        var idx = CCodeArr.indexOf(ecID);
                        evt["selectedIndex"] = idx;
                        // $(this).value=ecID
                        // $(this).val(ECCodeArr[i])
                        // $("#"+evt.id).val(ECCodeArr[i])
                    });
                }
                if (MCCodeArr.length > 0) {
                    $('.mobNoCode').each(function (i, evt) {
                        var ecID = MCCodeArr[i];
                        var idx = CCodeArr.indexOf(ecID);
                        evt["selectedIndex"] = idx;
                        // $(this).value=ecID
                        // $(this).val(ECCodeArr[i])
                        // $("#"+evt.id).val(ECCodeArr[i])
                    });
                }
                if (HCCodeArr.length > 0) {
                    $('.homeNoCode').each(function (i, evt) {
                        var ecID = HCCodeArr[i];
                        var idx = CCodeArr.indexOf(ecID);
                        evt["selectedIndex"] = idx;
                        // $(this).value=ecID
                        // $(this).val(ECCodeArr[i])
                        // $("#"+evt.id).val(ECCodeArr[i])
                    });
                }
                $(".addMobNo").click(function () {
                    multipleMobNo();
                });
                $(".addHomeNo").click(function () {
                    multipleHomeNo();
                });
                $(".addEmergencyNo").click(function () {
                    multipleEmergencyNo();
                });
                $("#EditedAddressDetails").html(OfficeAddArr.filter(function (add) { return SelectedUserProfile[0].Location == add.OfficePlace; })[0].OfficeFullAdd);
                $("#StaffFunctionEdit").val(SelectedUserProfile[0].Title);
                $("#StaffAffiliatesEdit").val(SelectedUserProfile[0].Affiliation);
                $("#PAddLineE").val(SelectedUserProfile[0].HAddLine);
                $("#PAddCityE").val(SelectedUserProfile[0].HAddCity);
                $("#PAddStateE").val(SelectedUserProfile[0].HAddState);
                $("#PAddPCodeE").val(SelectedUserProfile[0].HAddPCode);
                $("#PAddCountryE").val(SelectedUserProfile[0].HAddPCountry);
                $("#Eshortbio").val(SelectedUserProfile[0].ShortBio);
                $("#EIndustry").val(SelectedUserProfile[0].Industry);
                $("#ELanguage").val(SelectedUserProfile[0].Language);
                $("#ESDGCourse").val(SelectedUserProfile[0].SDGCourse);
                $("#ESoftwarExp").val(SelectedUserProfile[0].Software);
                $("#EMembership").val(SelectedUserProfile[0].Membership);
                $("#ESKnowledge").val(SelectedUserProfile[0].SpecialKnowledge);
                $("#citizenshipE").val(SelectedUserProfile[0].Citizen);
                $("#linkedInID").val(SelectedUserProfile[0].LinkedInID.Url);
                // $("#mobileno").val(SelectedUserProfile[0].PhoneNumber);
                $("#children").val(SelectedUserProfile[0].Child);
                $("#significantOther").val(SelectedUserProfile[0].SignOther);
                $("#USDDailyEdit").val(SelectedUserProfile[0].USDDaily);
                $("#USDHourlyEdit").val(SelectedUserProfile[0].USDHourly);
                $("#EURDailyEdit").val(SelectedUserProfile[0].EURDaily);
                $("#EURHourlyEdit").val(SelectedUserProfile[0].EURHourly);
                $("#assisstantName").val(SelectedUserProfile[0].Assistant);
                $("#personalmailID").val(SelectedUserProfile[0].UserPersonalMail);
                // $("#homeno").val(SelectedUserProfile[0].HomeNo);
                // $("#emergencyno").val(SelectedUserProfile[0].EmergencyNo);
                $("#workLocationDD").val(SelectedUserProfile[0].Location);
                $("#staffstatusDD").val(SelectedUserProfile[0].StaffStatus);
                $("#othercurrDD").val(SelectedUserProfile[0].OtherCurr);
                $("#ODailyEdit").val(SelectedUserProfile[0].OtherCurrDaily);
                $("#OHourlyEdit").val(SelectedUserProfile[0].OtherCurrHourly);
                if (SelectedUserProfile[0].StaffStatus == "Part-time") {
                    $("#workscheduleEdit").html("");
                    $("#workscheduleEdit").html("<div class=\"d-flex w-100\" id=\"workscheduleSec\"> <label>Work Schedule</label><div class=\"w-100\"><input type=\"text\" id=\"workScheduleE\" value=\"" + (SelectedUserProfile[0].WorkSchedule == "" || SelectedUserProfile[0].WorkSchedule == null ? "" : SelectedUserProfile[0].WorkSchedule) + "\"></div></div>");
                }
                else {
                    $("#workscheduleEdit").html("");
                    $("#workscheduleEdit").html("<div class=\"d-flex w-100 hide\" id=\"workscheduleSec\"> <label>Work Schedule</label><div class=\"w-100\"><input type=\"text\" id=\"workScheduleE\" value=\"" + (SelectedUserProfile[0].WorkSchedule == "" || SelectedUserProfile[0].WorkSchedule == null ? "" : SelectedUserProfile[0].WorkSchedule) + "\"></div></div>");
                }
                emailAddress = "i:0#.f|membership|" + SelectedUserProfile[0].AssistantMail.toLowerCase();
                divID = "peoplepickerText_TopSpan";
                SPClientPeoplePicker.SPClientPeoplePickerDict[divID].AddUnresolvedUser({
                    Key: emailAddress,
                    DisplayText: SelectedUserProfile[0].Assistant,
                    Email: SelectedUserProfile[0].AssistantMail.toLowerCase(),
                }, true);
                return [2 /*return*/];
        }
    });
}); };
var editsubmitFunction = function () { return __awaiter(_this, void 0, void 0, function () {
    var mobNumUpdate, homeNumUpdate, emergencyNumUpdate, mobNumbers, homeNumbers, emergencyNumbers, dispTitle, pickerDiv, peoplePicker, userInfo, profileID, loginName, profile, update, error_1;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0:
                mobNumUpdate = "";
                homeNumUpdate = "";
                emergencyNumUpdate = "";
                mobNumbers = document.querySelectorAll(".mobNumbers");
                homeNumbers = document.querySelectorAll(".homeNumbers");
                emergencyNumbers = document.querySelectorAll(".emergencyNumbers");
                mobNumbers.forEach(function (nums) {
                    mobNumUpdate += CCodeArr[nums.children[0]["options"].selectedIndex] + " - " + nums.children[1]["value"] + "^";
                });
                homeNumbers.forEach(function (nums) {
                    homeNumUpdate += CCodeArr[nums.children[0]["options"].selectedIndex] + " - " + nums.children[1]["value"] + "^";
                });
                emergencyNumbers.forEach(function (nums) {
                    emergencyNumUpdate += CCodeArr[nums.children[0]["options"].selectedIndex] + " - " + nums.children[1]["value"] + "^";
                });
                if (bioAttachArr.length > 0) {
                    bioAttachArr.map(function (filedata) {
                        sp.web.folders
                            .add("/sites/StaffDirectory/BiographyDocument/" + selectedUsermail)
                            .then(function (data) {
                            sp.web
                                .getFolderByServerRelativeUrl(data.data.ServerRelativeUrl)
                                .files.add(filedata.name, filedata.content, true);
                        });
                    });
                }
                dispTitle = "APickerField";
                pickerDiv = $("[id$='peoplepickerText'][title='" + dispTitle + "']");
                peoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict;
                userInfo = peoplePicker.peoplepickerText_TopSpan.GetAllUserInfo();
                profileID = 0;
                if (!(userInfo.length > 0)) return [3 /*break*/, 2];
                loginName = userInfo[0].Key.split("|")[2];
                return [4 /*yield*/, sp.web.siteUsers.getByEmail(loginName).get()];
            case 1:
                profile = _a.sent();
                profileID = profile.Id;
                _a.label = 2;
            case 2:
                _a.trys.push([2, 4, , 5]);
                return [4 /*yield*/, sp.web
                        .getList(listUrl + "StaffDirectory")
                        .items.getById(ItemID)
                        .update({
                        Title: "SDG User Info",
                        PersonalEmail: $("#personalmailID").val(),
                        MobileNo: mobNumUpdate,
                        HomeNo: homeNumUpdate,
                        EmergencyNo: emergencyNumUpdate,
                        HomeAddLine: $("#PAddLineE").val(),
                        HomeAddCity: $("#PAddCityE").val(),
                        HomeAddState: $("#PAddStateE").val(),
                        HomeAddPCode: $("#PAddPCodeE").val(),
                        HomeAddCountry: $("#PAddCountryE").val(),
                        IndustryExp: $("#EIndustry").val(),
                        LanguageExp: $("#ELanguage").val(),
                        SDGCourses: $("#ESDGCourse").val(),
                        SoftwareExp: $("#ESoftwarExp").val(),
                        Membership: $("#EMembership").val(),
                        SpecialKnowledge: $("#ESKnowledge").val(),
                        Citizenship: $("#citizenshipE").val(),
                        ShortBio: $("#Eshortbio").val(),
                        USDDailyRate: $("#USDDailyEdit").val(),
                        USDHourlyRate: $("#USDHourlyEdit").val(),
                        EURDailyRate: $("#EURDailyEdit").val(),
                        EURHourlyRate: $("#EURHourlyEdit").val(),
                        OtherCurrency: $("#othercurrDD").val(),
                        ODailyRate: $("#ODailyEdit").val(),
                        OHourlyRate: $("#OHourlyEdit").val(),
                        EffectiveDate: $("#EffectiveDateEdit").val(),
                        signother: $("#significantOther").val(),
                        children: $("#children").val(),
                        WorkingSchedule: $("#workSchedule").val(),
                        SDGOffice: $("#workLocationDD").val(),
                        StaffStatus: $("#staffstatusDD").val(),
                        LinkedInLink: {
                            "__metadata": { type: "SP.FieldUrlValue" },
                            Description: "LinkedIn",
                            Url: $("#linkedInID").val()
                        },
                        stafffunction: $("#StaffFunctionEdit").val(),
                        SDGAffiliation: $("#StaffAffiliatesEdit").val(),
                        AssistantId: profileID,
                    })];
            case 3:
                update = _a.sent();
                location.reload();
                return [3 /*break*/, 5];
            case 4:
                error_1 = _a.sent();
                console.log(error_1);
                return [3 /*break*/, 5];
            case 5: return [2 /*return*/];
        }
    });
}); };
var editcancelFunction = function () {
    var viewDir = document.querySelector(".view-directory");
    var editDir = document.querySelector(".edit-directory");
    var editbtn = document.querySelector(".btn-edit");
    viewDir.classList.remove("hide");
    editDir.classList.add("hide");
    editbtn.classList.remove("hide");
    //  $('#peoplepickerText').children().remove();
    // sp-peoplepicker-editorInput
};
var useravailabilityDetails = function () { return __awaiter(_this, void 0, void 0, function () {
    var availTableHtml, options, userAvailTable;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0:
                console.log(SelectedUserProfile[0].Name);
                return [4 /*yield*/, sp.web
                        .getList(listUrl + "SDGAvailability")
                        .items.select("*", "UserName/EMail", "UserName/Id").expand("UserName").filter("UserName/EMail eq '" + SelectedUserProfile[0].Usermail + "'")
                        .getAll()];
            case 1:
                availList = _a.sent();
                console.log(availList);
                availTableHtml = "";
                availList.forEach(function (avail) {
                    availTableHtml += "<tr><td>" + avail.Project + "</td><td>" + new Date(avail.StartDate).toLocaleDateString() + "</td><td>" + new Date(avail.EndDate).toLocaleDateString() + "</td><td>" + avail.Percentage + "%</td><td>" + avail.Comments + "</td><td><div class=\"d-flex\"><div class=\"action-btn action-edit\" data-toggle=\"modal\" data-target=\"#addprojectmodal\" data-id=\"" + avail.ID + "\" id=\"editProjectAvailability\"></div><div class=\"action-btn action-delete\" data-id=\"" + avail.ID + "\" id=\"deleteProjectAvailability\"> </div></div></td></tr>";
                });
                $("#UserAvailabilityTbody").html("");
                $("#UserAvailabilityTbody").html(availTableHtml);
                options = {
                    order: [[0, "asc"]],
                    destroy: true,
                };
                userAvailTable = $("#UserAvailabilityTable").DataTable(options);
                $('.usernametag').on('click', function () {
                    userAvailTable.destroy();
                });
                return [2 /*return*/];
        }
    });
}); };
function removeSelectedfile(filename) {
    for (var i = 0; i < bioAttachArr.length; i++) {
        if (bioAttachArr[i].name == filename) {
            ///filesQuantity[i].remove();
            bioAttachArr.splice(i, 1);
            break;
        }
    }
}
var removeAvailProject = function (ID) {
    sp.web.getList(listUrl + "SDGAvailability").items.getById(ID).delete();
};
var multipleMobNo = function () {
    $("#mobileNoSec").append("<div class=\"d-flex mobNumbers\"><select class=\"mobNoCode\">" + CCodeHtml + "</select><input type=\"number\" class=\"mobNo\" id=\"mobileno1\"/><span class=\"removeMobNo remove-icon\"></span></div>");
};
var multipleHomeNo = function () {
    $("#homeNoSec").append("<div class=\"d-flex homeNumbers\"><select class=\"homeNoCode\">" + CCodeHtml + "</select><input type=\"number\" class=\"mobNo\" id=\"homeno1\"/><span class=\"removeHomeNo remove-icon\"></span></div>");
};
var multipleEmergencyNo = function () {
    $("#emergencyNoSec").append("<div class=\"d-flex emergencyNumbers\"><select class=\"emergencyNoCode\">" + CCodeHtml + "</select><input type=\"number\" class=\"mobNo\" id=\"emergencyno1\"/><span class=\"removeHomeNo remove-icon\"></span></div>");
};
var fillEditSection = function (ID) {
    //  sp.web
    // .getList(listUrl + "SDGAvailability").items.getById(ID).get().then((item:any)=>{
    var editedData = availList.filter(function (e) { return e.Id == parseInt(ID); });
    if (editedData.length > 0) {
        // var finalmonth: any = "";
        // var dd = new Date(SelectedUserProfile[0].EffectiveDate).getDate();
        // var mm = new Date(SelectedUserProfile[0].EffectiveDate).getMonth() + 1;
        // mm < 10 ? (finalmonth = "0" + mm) : (finalmonth = mm);
        // var yyyy = new Date(SelectedUserProfile[0].EffectiveDate).getFullYear();
        // var dateformat = yyyy + "-" + finalmonth + "-" + dd;
        // $("#EffectiveDateEdit").val(dateformat);
        var Sfinalmonth = "";
        var Sfinalday = "";
        var Sdd = new Date(editedData[0].StartDate).getDate();
        Sdd < 10 ? (Sfinalday = "0" + Sdd) : (Sfinalday = Sdd);
        var Smm = new Date(editedData[0].StartDate).getMonth() + 1;
        Smm < 10 ? (Sfinalmonth = "0" + Smm) : (Sfinalmonth = Smm);
        var Syyyy = new Date(editedData[0].StartDate).getFullYear();
        var Sdateformat = Syyyy + "-" + Sfinalmonth + "-" + Sfinalday;
        var Efinalmonth = "";
        var Edd = new Date(editedData[0].EndDate).getDate();
        var Emm = new Date(editedData[0].EndDate).getMonth() + 1;
        Emm < 10 ? (Efinalmonth = "0" + Smm) : (Efinalmonth = Emm);
        var Eyyyy = new Date(editedData[0].EndDate).getFullYear();
        var Edateformat = Eyyyy + "-" + Efinalmonth + "-" + Edd;
        $("#projectName").val(editedData[0].Project);
        $("#projectStartDate").val(Sdateformat);
        $("#projectEndDate").val(Edateformat);
        $("#projectPercent").val(editedData[0].Percentage);
        $("#practiceAreaDD").val(editedData[0].ProjectArea);
        $("#client").val(editedData[0].Client);
        $("#projectCode").val(editedData[0].ProjectCode);
        $("#ProjectLocation").val(editedData[0].ProjectLocation);
        $("#projectAvailNotes").val(editedData[0].Notes);
        $("#Projectcomments").val(editedData[0].Comments);
    }
    // })
};
var availSubmitFunc = function () { return __awaiter(_this, void 0, void 0, function () {
    var bwArray, isAllSuccess, sDate, eDate, startd, endd, newend, newDate, enteredPercentage, _loop_1, dayValue, correlationPercentage, i, state_1, ProjectPercent, submitProject;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0:
                bwArray = [];
                isAllSuccess = true;
                console.log(availList);
                sDate = $("#projectStartDate").val();
                eDate = $("#projectEndDate").val();
                startd = new Date(sDate);
                endd = new Date(eDate);
                newend = endd.setDate(endd.getDate() + 1);
                endd = new Date(newend);
                while (startd < endd) {
                    bwArray.push(new Date(startd).toLocaleDateString());
                    console.log(startd); // ISO Date format          
                    newDate = startd.setDate(startd.getDate() + 1);
                    startd = new Date(newDate);
                }
                enteredPercentage = parseInt($("#projectPercent").val());
                _loop_1 = function (i) {
                    var datearr = bwArray[i];
                    var filteredData = [];
                    availList.filter(function (data) {
                        if (new Date(data.StartDate).toLocaleDateString() <= datearr && new Date(data.EndDate).toLocaleDateString() >= datearr) {
                            filteredData.push(data);
                        }
                    });
                    dayValue = filteredData.reduce(function (n, _a) {
                        var Percentage = _a.Percentage;
                        return n + parseInt(Percentage);
                    }, 0);
                    correlationPercentage = 100 - parseInt(dayValue);
                    if (enteredPercentage <= correlationPercentage) {
                        console.log('Suceess');
                        isAllSuccess = true;
                    }
                    else {
                        alert("Not able to add ur percentage in this date : " + datearr);
                        isAllSuccess = false;
                        return "break";
                    }
                };
                // bwArray.map((datearr)=>{
                for (i = 0; i < bwArray.length; i++) {
                    state_1 = _loop_1(i);
                    if (state_1 === "break")
                        break;
                }
                ProjectPercent = 0;
                if (!isAllSuccess) return [3 /*break*/, 2];
                return [4 /*yield*/, sp.web.getList(listUrl + "SDGAvailability").items.add({
                        UserNameId: SelectedUserProfile[0].UserId,
                        Project: $("#projectName").val(),
                        StartDate: $("#projectStartDate").val(),
                        EndDate: $("#projectEndDate").val(),
                        Percentage: enteredPercentage.toString(),
                        ProjectArea: $("#practiceAreaDD").val(),
                        Client: $("#client").val(),
                        ProjectCode: $("#projectCode").val(),
                        ProjectLocation: $("#ProjectLocation").val(),
                        Notes: $("#projectAvailNotes").val(),
                        Comments: $("#Projectcomments").val()
                    })];
            case 1:
                submitProject = _a.sent();
                $("#projectName").val("");
                $("#projectStartDate").val("");
                $("#projectEndDate").val("");
                $("#projectPercent").val("");
                $("#practiceAreaDD").val("");
                $("#client").val("");
                $("#projectCode").val("");
                $("#ProjectLocation").val("");
                $("#projectAvailNotes").val("");
                $("#Projectcomments").val("");
                location.reload();
                return [3 /*break*/, 3];
            case 2:
                alert("Available Percentage is:" + correlationPercentage);
                _a.label = 3;
            case 3: return [2 /*return*/];
        }
    });
}); };
var availUpdateFunc = function () {
    var bwArray = [];
    var isAllSuccess = true;
    var sDate = $("#projectStartDate").val();
    var eDate = $("#projectEndDate").val();
    var startd = new Date(sDate);
    var endd = new Date(eDate);
    var newend = endd.setDate(endd.getDate() + 1);
    endd = new Date(newend);
    while (startd < endd) {
        bwArray.push(new Date(startd).toLocaleDateString());
        console.log(startd); // ISO Date format          
        var newDate = startd.setDate(startd.getDate() + 1);
        startd = new Date(newDate);
    }
    var enteredPercentage = parseInt($("#projectPercent").val());
    var _loop_2 = function (i) {
        var datearr = bwArray[i];
        var filteredData = [];
        availList.filter(function (data) {
            if (new Date(data.StartDate).toLocaleDateString() <= datearr && new Date(data.EndDate).toLocaleDateString() >= datearr) {
                filteredData.push(data);
            }
        });
        dayValue = filteredData.reduce(function (n, _a) {
            var Percentage = _a.Percentage;
            return n + parseInt(Percentage);
        }, 0);
        correlationPercentage = 100 - parseInt(dayValue);
        if (enteredPercentage <= correlationPercentage) {
            console.log('Suceess');
            isAllSuccess = true;
        }
        else {
            alert("Not able to add ur percentage in this date : " + datearr);
            isAllSuccess = false;
            return "break";
        }
    };
    var dayValue, correlationPercentage;
    // bwArray.map((datearr)=>{
    for (var i = 0; i < bwArray.length; i++) {
        var state_2 = _loop_2(i);
        if (state_2 === "break")
            break;
    }
    var ProjectPercent = 0;
    $("#projectPercent").val() == "" ? ProjectPercent = 0 : ProjectPercent = parseInt($("#projectPercent").val());
    var updateProject = sp.web
        .getList(listUrl + "SDGAvailability").items.getById(AvailEditID).update({
        UserName: SelectedUserProfile[0].Name,
        Project: $("#projectName").val(),
        StartDate: $("#projectStartDate").val(),
        EndDate: $("#projectEndDate").val(),
        Percentage: ProjectPercent.toString(),
        ProjectArea: $("#practiceAreaDD").val(),
        Client: $("#client").val(),
        ProjectCode: $("#projectCode").val(),
        ProjectLocation: $("#ProjectLocation").val(),
        Notes: $("#projectAvailNotes").val(),
        Comments: $("#Projectcomments").val()
    });
    $("#projectName").val("");
    $("#projectStartDate").val("");
    $("#projectEndDate").val("");
    $("#projectPercent").val("");
    $("#practiceAreaDD").val("");
    $("#client").val("");
    $("#projectCode").val("");
    $("#ProjectLocation").val("");
    $("#projectAvailNotes").val("");
    $("#Projectcomments").val("");
    AvailEditID = 0;
    AvailEditFlag = false;
    location.reload();
};
//# sourceMappingURL=StaffdirectoryWebPart.js.map