'use strict';
const {request}=require('./http.cjs');
const {URL}=require('node:url');
const {normalizeCredential}=require('./credentials.cjs');
const {requireCondition:check}=require('./errors.cjs');
const issueFields='id identifier title description url team { id } state { id name type } labels(first:100) { nodes { id } }';
const attachmentFields='id title url metadata issue { id }';
class LinearClient{
  constructor({apiKey,workspaceId,teamId,apiUrl='https://api.linear.app/graphql',transport=request,simulation=false,maxPages=20}){
    check(apiKey&&workspaceId&&teamId,'LINEAR_CONFIGURATION_REQUIRED');
    check(apiUrl==='https://api.linear.app/graphql'||simulation===true,'LINEAR_ENDPOINT_FORBIDDEN');
    Object.assign(this,{apiKey,workspaceId,teamId,apiUrl,transport,simulation,maxPages});this.checked=false;
    this.apiKey=normalizeCredential(apiKey,'linear');
  }
  async gql(operation,query,variables={}){
    const r=await this.transport(this.apiUrl,{method:'POST',headers:{Authorization:this.apiKey,'Content-Type':'application/json'},body:Buffer.from(JSON.stringify({operationName:operation,query,variables})),maxBytes:2097152,allowHttp:this.simulation});
    check(r.status!==401&&r.status!==403,'LINEAR_AUTH_FAILED');check(r.status!==429,'LINEAR_RATE_LIMITED');check(r.status===200,'LINEAR_HTTP_FAILED');
    let json;try{json=JSON.parse(r.bytes.toString());}catch{check(false,'LINEAR_INVALID_JSON');}
    check(!json.errors?.length&&json.data,'LINEAR_GRAPHQL_FAILED');return json.data;
  }
  async assertScope(){
    const d=await this.gql('CartorioScope','query CartorioScope { organization { id name } }');
    check(d.organization?.id===this.workspaceId,'LINEAR_WORKSPACE_FORBIDDEN');this.checked=true;return {id:d.organization.id,name:d.organization.name};
  }
  async readyIssues({labelId,readyStateIds,maxIssues=200}){
    check(this.checked,'LINEAR_SCOPE_NOT_VERIFIED');const all=[];let after=null;
    for(let page=0;page<this.maxPages;page++){
      const d=await this.gql('CartorioReady','query CartorioReady($filter:IssueFilter!,$after:String){ issues(filter:$filter,first:50,after:$after){ nodes { id } pageInfo { hasNextPage endCursor } } }',{filter:{team:{id:{eq:this.teamId}},labels:{some:{id:{eq:labelId}}},state:{id:{in:readyStateIds}}},after});
      all.push(...d.issues.nodes.map(i=>i.id));if(all.length>=maxIssues)return [...new Set(all)].slice(0,maxIssues);if(!d.issues.pageInfo.hasNextPage)return [...new Set(all)];
      check(d.issues.pageInfo.endCursor&&d.issues.pageInfo.endCursor!==after,'LINEAR_PAGINATION_INVALID');after=d.issues.pageInfo.endCursor;
    }check(false,'LINEAR_PAGINATION_LIMIT');
  }
  async issue(id){
    check(this.checked,'LINEAR_SCOPE_NOT_VERIFIED');
    const d=await this.gql('CartorioIssue',`query CartorioIssue($id:String!){ issue(id:$id){ ${issueFields} attachments(first:100){ nodes { ${attachmentFields} } pageInfo { hasNextPage endCursor } } } }`,{id});
    const i=d.issue;check(i?.id===id&&i.team?.id===this.teamId,'LINEAR_ISSUE_SCOPE_MISMATCH');
    check(!i.attachments.pageInfo.hasNextPage,'LINEAR_ATTACHMENT_LIMIT');
    check(i.attachments.nodes.every(a=>a.issue?.id===i.id),'LINEAR_ATTACHMENT_ISSUE_MISMATCH');return i;
  }
  async issueStatus(id){
    check(this.checked,'LINEAR_SCOPE_NOT_VERIFIED');
    const d=await this.gql('CartorioIssueStatus','query CartorioIssueStatus($id:String!){issue(id:$id){id team{id} state{id name type} labels(first:100){nodes{id}}}}',{id});
    check(d.issue?.id===id&&d.issue.team?.id===this.teamId,'LINEAR_ISSUE_SCOPE_MISMATCH');return d.issue;
  }
  async document(id){
    check(this.checked,'LINEAR_SCOPE_NOT_VERIFIED');
    const d=await this.gql('CartorioTemplateDocument','query CartorioTemplateDocument($id:String!){ document(id:$id){ id title url updatedAt archivedAt trashed content team { id } } }',{id});
    check(d.document?.id===id&&d.document.team?.id===this.teamId,'TEMPLATE_DOCUMENT_TEAM_MISMATCH');
    check(!d.document.archivedAt&&!d.document.trashed,'TEMPLATE_DOCUMENT_UNAVAILABLE');return d.document;
  }
  checkAsset(url,upload=false){
    const u=new URL(url);check(!u.username&&!u.password,'ASSET_URL_FORBIDDEN');
    if(this.simulation){check(u.origin===new URL(this.apiUrl).origin,'ASSET_URL_FORBIDDEN');return;}
    check(u.protocol==='https:'&&!u.port&&((!upload&&u.hostname==='uploads.linear.app')||(upload&&(u.hostname==='uploads.linear.app'||u.hostname==='storage.googleapis.com'||/^[a-z0-9.-]+\.storage\.googleapis\.com$/.test(u.hostname)))),'ASSET_URL_FORBIDDEN');
  }
  async download(url,maxBytes=20971520){
    this.checkAsset(url);const r=await this.transport(url,{headers:{Authorization:this.apiKey},maxBytes,allowHttp:this.simulation});
    check(r.status===200,'LINEAR_DOWNLOAD_FAILED');return r.bytes;
  }
  async allocate(filename,size,contentType){
    check(this.checked,'LINEAR_SCOPE_NOT_VERIFIED');
    const d=await this.gql('CartorioUpload','mutation CartorioUpload($contentType:String!,$filename:String!,$size:Int!){ fileUpload(contentType:$contentType,filename:$filename,size:$size){ success uploadFile { uploadUrl assetUrl headers { key value } } } }',{contentType,filename,size});
    check(d.fileUpload?.success&&d.fileUpload.uploadFile,'LINEAR_UPLOAD_ALLOCATION_FAILED');const f=d.fileUpload.uploadFile;this.checkAsset(f.assetUrl);this.checkAsset(f.uploadUrl,true);return f;
  }
  async put(file,bytes,contentType){
    this.checkAsset(file.uploadUrl,true);const headers={'Content-Type':contentType};
    for(const h of file.headers){check(!/^(authorization|host|cookie)$/i.test(h.key),'UNSAFE_UPLOAD_HEADER');headers[h.key]=h.value;}
    const r=await this.transport(file.uploadUrl,{method:'PUT',headers,body:bytes,maxBytes:1048576,allowHttp:this.simulation});check(r.status>=200&&r.status<300,'LINEAR_UPLOAD_FAILED');
  }
  async attach(issueId,assetUrl,title,metadata){
    this.checkAsset(assetUrl);await this.issue(issueId);
    const d=await this.gql('CartorioAttach',`mutation CartorioAttach($input:AttachmentCreateInput!){ attachmentCreate(input:$input){ success attachment { ${attachmentFields} } } }`,{input:{issueId,url:assetUrl,title,metadata}});
    check(d.attachmentCreate?.success&&d.attachmentCreate.attachment?.issue?.id===issueId,'LINEAR_ATTACHMENT_CREATE_FAILED');return d.attachmentCreate.attachment;
  }
  async comment(issueId,body,marker){
    const d=await this.gql('CartorioComments','query CartorioComments($id:String!){ issue(id:$id){ id comments(first:100){ nodes { id body } pageInfo { hasNextPage } } } }',{id:issueId});
    check(d.issue?.id===issueId&&!d.issue.comments.pageInfo.hasNextPage,'LINEAR_COMMENT_LIMIT');
    if(d.issue.comments.nodes.some(c=>c.body.includes(marker)))return;
    const result=await this.gql('CartorioComment','mutation CartorioComment($input:CommentCreateInput!){ commentCreate(input:$input){ success comment { id } } }',{input:{issueId,body}});
    check(result.commentCreate?.success,'LINEAR_COMMENT_FAILED');
  }
  async review(issueId,stateId){
    await this.issue(issueId);
    const d=await this.gql('CartorioReview','mutation CartorioReview($id:String!,$input:IssueUpdateInput!){ issueUpdate(id:$id,input:$input){ success issue { id state { id } } } }',{id:issueId,input:{stateId}});
    check(d.issueUpdate?.success&&d.issueUpdate.issue?.id===issueId&&d.issueUpdate.issue.state.id===stateId,'LINEAR_REVIEW_UPDATE_FAILED');
  }
}
module.exports={LinearClient};
