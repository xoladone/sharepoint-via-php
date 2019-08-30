<?php
/**
 * This is the general handler for sharepoint.
 * 
 * Sharepoint is accessed via the graph api v1. 
 * This tutorial was helpful: https://medium.com/@anoopt/access-sharepoint-data-using-postman-eec5965400f2
 * Great playground: https://developer.microsoft.com/en-us/graph/graph-explorer#
 * 
 * @filesource 
 * @version 0.0.2
 * @author Daniel Jansen
 * @copyright 2019 Fischer Akkumulatorentechnik GmbH
 * 
 */


class sharepoint {
  // variables for sharepoint
  private $grant_type;
  private $client_id;
  private $client_secret;
  private $scope;
  private $directory_id;
  private $url;
  protected $host_name;
  public $token;
  
 /**
  * Short Description1
  * 
  * Long Description2
  * @param mixed[] $items Array structure to count the elements of.
  * @return boolean TRUE | False Returns the number of elements.
  */  
  public function __construct() {

    // set credentials
    $this->grant_type       = 'client_credentials';
    $this->client_id        = SP_CLIENTID; // This constant is to change by you.
    $this->client_secret    = SP_CLIENTSECRET; // This constant is to change by you.
    $this->scope            = 'https://graph.microsoft.com/.default';
    $this->directory_id     = SP_DIRECTORYID; // This constant is to change by you.
    $this->host_name        = SP_HOSTNAME; // This constant is to change by you.
    $this->url['token']     = 'https://login.microsoftonline.com/'.$this->directory_id.'/oauth2/v2.0/token';
    $this->url['sites']     = 'https://graph.microsoft.com/v1.0/sites/';
    
    // get token for operations
    $this->token            = $this->Token_get();
  }
  
  /**
   * gets the graphapi token 
   * 
   * Calls the ms graphapi signin url and gets a token with the credentials set in __construct()
   * @see sharepoint::__construct()
   * @return mixed[] JSON with the token
   */   
  public function Token_get() {
    $curl = curl_init();
    curl_setopt_array($curl, array(
      CURLOPT_URL => $this->url['token'],
      CURLOPT_RETURNTRANSFER => true,
      CURLOPT_ENCODING => "",
      CURLOPT_MAXREDIRS => 10,
      CURLOPT_TIMEOUT => 30,
      CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
      CURLOPT_CUSTOMREQUEST => "POST",
      CURLOPT_POSTFIELDS => "grant_type=$this->grant_type&client_id=$this->client_id&client_secret=$this->client_secret&scope=$this->scope",
      CURLOPT_HTTPHEADER => array(
        "Accept: */*",
        "Accept-Encoding: gzip, deflate",
        "Cache-Control: no-cache",
        "Connection: keep-alive",
      ),
    ));
    $response = curl_exec($curl);
    $err = curl_error($curl);
    $errNo = curl_errno($curl);
    curl_close($curl);
    if ($err) {
      return array ('error' => $err, 'error_number' => $errNo);
    } else {
      $array = json_decode($response, true);
      $array['expires_timestamp'] = time() + 3540; // sets expire-timestamp to 59 minutes in the future
      return $array;
    }
  }
  
  /**
   * Gets the id of a folder.
   * 
   * Returns a json with the id of the relative specified folder or file.
   * @see sharepoint::SiteId_get()
   * @param string $siteId The id of the subsite
   * @param string $relativePath The relative path of the folder. eg 'folder1/folder2', or 'folder/file.txt'
   * @return json ID of the folder or error.
   */  
  public function DriveItem_get($siteId, $relativePath) {
    $url          = $this->url['sites'].$siteId .'/drive/root:/'.$relativePath;
    $postfields   = '';
    $method       = 'GET';
    $contentType  = '';
    if (time() > $this->token['expires_timestamp']) {
      $this->Token_get();
    }
    $return = $this->curl($url, $postfields, $method, $contentType);
    return $return;
  }
  
  /**
   * Uploads a file. Just works up to 4 megabyte.
   * 
   * Uploads a binary file to a specified relative path. This upload-method is limited up to 4 MB by the ms graphapi.
   * @see sharepoint::SiteId_get()
   * @param string $siteId The id of the subsite
   * @param string $relativePath eg '/documents/data'
   * @param string $fileName The name of the new file
   * @param blob $fileBinary use s.th. like fread(fopen($filename, "rb"), filesize($filename));
   * @return json Returns a json with data like the new file-id or an error
   */  
  public function File_upload($siteId, $relativePath, $fileName, $fileBinary) {
    $url          = $this->url['sites'].$siteId .'/drive/root:/'.$relativePath.$fileName.':/content';
    $postfields   = $fileBinary;
    $method       = 'PUT';
    $contentType  = 'application/x-binary';
    if (time() > $this->token['expires_timestamp']) {
      $this->Token_get();
    }
    $return = $this->curl($url, $postfields, $method, $contentType);
    return $return;
  }

  /**
   * Updates the fields (values of columns) of a file
   * 
   * To update the fields of an item.
   * Warning: The field names can differ from what is seen on the Website. Its allways the first name. 
   * @see sharepoint::SiteId_get()
   * @param string $siteId The id of the subsite.
   * @param string $itemId The id of the item (file).
   * @param string $fields A Json-String with key value pairs, eg {"Colname1":"Colvalue1", "Colname2":"Colvalue2"}
   * @return json Returns a json with field data or errors
   */  
  public function FileFields_update($siteId, $itemId, $fields) {
    $url          = $this->url['sites'].$siteId .'/drive/items/'.$itemId.'/listItem/fields';
    $postfields   = $fields;
    $method       = 'PATCH';
    $contentType  = 'application/json';
    if (time() > $this->token['expires_timestamp']) {
      $this->Token_get();
    }
    $return = $this->curl($url, $postfields, $method, $contentType);
    return $return;
  }
  
  /**
   * Creates a new folder in a directory
   * 
   * Creates a new folder in a directoy given with the $parentItemId
   * @see sharepoint::DriveItem_get()
   * @see sharepoint::SiteId_get()
   * @param string $siteId The id of the subsite
   * @param string $parentItemId The id of the parent folder
   * @param string $folderName The name of the new folder
   * @return boolean TRUE | False Returns the number of elements.
   */  
  public function Folder_create($siteId, $parentItemId, $folderName) {
    $url          = $this->url['sites'].$siteId .'/drive/items/'.$parentItemId.'/children';
    $postfields   = '{"name": "'.$folderName.'", "folder":{}}';
    $method       = 'POST';
    $contentType  = 'application/json';
    if (time() > $this->token['expires_timestamp']) {
      $this->Token_get();
    }
    $return = $this->curl($url, $postfields, $method, $contentType);
    return $return;
  }
  
  /**
   * Gets the id of a subsite 
   * 
   * Return a json with the id of the subsite with the relative path
   * @param string $relativePath string with the relative path of the subsite eg 'nameOfTheSubsite'
   * @return mixed[] Returns a json with the id of the subsite
   */  
  public function SiteId_get($relativePath) {
    $url          = $this->url['sites'].$this->host_name .':/'.$relativePath;
    $postfields   = '';
    $method       = 'GET';
    $contentType  = '';
    if (time() > $this->token['expires_timestamp']) {
      $this->Token_get();
    }
    $return = $this->curl($url, $postfields, $method, $contentType);
    return $return;
  }
  
  /**
   * Short Description1
   * 
   * Long Description2
   * @see MyClass::$items
   * @param mixed[] $items Array structure to count the elements of.
   * @return boolean TRUE | False Returns the number of elements.
   */  
  public function ListId_get($siteId) {
    $url          = $this->url['sites'].$siteId.'/lists';
    $postfields   = '';
    $method       = 'GET';
    $contentType  = '';
    if (time() > $this->token['expires_timestamp']) {
      $this->Token_get();
    }
    $return = $this->curl($url, $postfields, $method, $contentType);
    return $return;
  }
  
  /**
   * Short Description1
   * 
   * Long Description2
   * @see MyClass::$items
   * @param mixed[] $items Array structure to count the elements of.
   * @return boolean TRUE | False Returns the number of elements.
   */  
  public function ListItems_get($siteId, $listId, $query='') {
    $url          = $this->url['sites'].$siteId.'/lists/'.$listId.'/items'.$query;
    $postfields   = '';
    $method       = 'GET';
    $contentType  = '';
    if (time() > $this->token['expires_timestamp']) {
      $this->Token_get();
    }
    $return = $this->curl($url, $postfields, $method, $contentType);
    return $return;
  }

  /**
   * creates a text-file in a specified folder
   * 
   * Creates a text-file in the subsite an folder specified by their ids.
   * The filename becomes url-encoded within this function.
   * @param string $siteID The id of the subsite
   * @param string $parentID The id of the folder where to create the file
   * @param string $fileName  The name of the file.
   * @param string $fileContent The Content of the file.
   * @return boolean TRUE | False Returns the number of elements.
   */  
  protected function TextFile_put ($siteID, $parentID, $fileName, $fileContent) {
    $fileNameEnc = strval(rawurlencode($fileName));
    $url          = $this->url['sites'].$siteID .'/drive/items/'.$parentID.':/'.$fileNameEnc.':/content';
    $postfields   = $fileContent;
    $method       = 'PUT';
    $contentType  = 'text/plain';
    if (time() > $this->token['expires_timestamp']) {
      $this->Token_get();
    }
    $return = $this->curl($url, $postfields, $method, $contentType);
  }

  /**
   * curl function to call the rest api of ms graphapi
   * 
   * Long Description2
   * @param string $url The Endpoint of the restapi
   * @param string $postfields The Content for the request body. eg JSon-text for folder create or binary for fileupload
   * @param string $method GET, POST, PUT, ...
   * @param string $contentType String eg 'application/json' or 'application/x-binary' default is ''
   * @return boolean TRUE | False Returns the number of elements.
   */  
  private function curl($url, $postfields, $method='GET', $contentType='') {
    $curl = curl_init();
    curl_setopt_array($curl, array(
      CURLOPT_URL => $url,
      CURLOPT_RETURNTRANSFER => true,
      CURLOPT_ENCODING => "",
      CURLOPT_MAXREDIRS => 10,
      CURLOPT_TIMEOUT => 30,
      CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
      CURLOPT_CUSTOMREQUEST => $method,
      CURLOPT_POSTFIELDS => $postfields,
      CURLOPT_HTTPHEADER => array(
        "Accept: */*",
        "Content-Type: ".$contentType,
        "Accept-Encoding: gzip, deflate",
        "Cache-Control: no-cache",
        "Connection: keep-alive",
        "Authorization:Bearer " . $this->token['access_token']
      ),
    ));
    $response = curl_exec($curl);
    $err = curl_error($curl);
    $errNo = curl_errno($curl);
    curl_close($curl);
    if ($err) {
      return array ('error' => $err, 'error_number' => $errNo);
    } else {
      return json_decode($response, true);
    }
  }  
}
