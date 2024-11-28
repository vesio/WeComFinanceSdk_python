import base64
import ctypes
import json
import os
import time
import hashlib
from ctypes import c_char_p, c_int, c_ulonglong, POINTER, create_string_buffer

from Crypto.Cipher import PKCS1_v1_5
from Crypto.PublicKey import RSA

if os.name == "nt":
    sdk_dll = ctypes.CDLL("./WeWorkFinanceSdk.dll")
elif os.name == 'posix':
    sdk_dll = ctypes.CDLL("./libWeWorkFinanceSdk_C.so")
else:
    raise NotImplementedError("Unsupported OS")

const_prikey_pem_path = "./prikey.pem"

# 定义SDK结构体
class Slice(ctypes.Structure):
    _fields_ = [("buf", c_char_p),
                ("len", c_int)]

class MediaData(ctypes.Structure):
    _fields_ = [("outindexbuf", c_char_p),
                ("out_len", c_int),
                # data 是二进制文件所以必须用 c_void_p , 不能用 c_char_p
                ("data", ctypes.c_void_p),
                ("data_len", c_int),
                ("is_finish", c_int)]

# 定义SDK函数原型
sdk_dll.NewSdk.restype = ctypes.c_void_p

"""	/**
	 * 初始化函数
	 * Return值=0表示该API调用成功
	 * 
	 * @param [in]  sdk			NewSdk返回的sdk指针
	 * @param [in]  corpid      调用企业的企业id，例如：wwd08c8exxxx5ab44d，可以在企业微信管理端--我的企业--企业信息查看
	 * @param [in]  secret		聊天内容存档的Secret，可以在企业微信管理端--管理工具--聊天内容存档查看
	 *						
	 *
	 * @return 返回是否初始化成功
	 *      0   - 成功
	 *      !=0 - 失败
	 */
	 int Init(WeWorkFinanceSdk_t* sdk, const char* corpid, const char* secret);
"""
sdk_dll.Init.argtypes = [ctypes.c_void_p, c_char_p, c_char_p]
sdk_dll.Init.restype = c_int

"""
	/**
	 * 拉取聊天记录函数
	 * Return值=0表示该API调用成功
	 * 
	 *
	 * @param [in]  sdk				NewSdk返回的sdk指针
	 * @param [in]  seq				从指定的seq开始拉取消息，注意的是返回的消息从seq+1开始返回，seq为之前接口返回的最大seq值。首次使用请使用seq:0
	 * @param [in]  limit			一次拉取的消息条数，最大值1000条，超过1000条会返回错误
	 * @param [in]  proxy			使用代理的请求，需要传入代理的链接。如：socks5://10.0.0.1:8081 或者 http://10.0.0.1:8081
	 * @param [in]  passwd			代理账号密码，需要传入代理的账号密码。如 user_name:passwd_123
	 * @param [in]  timeout			超时时间，单位秒
	 * @param [out] chatDatas		返回本次拉取消息的数据，slice结构体.内容包括errcode/errmsg，以及每条消息内容。示例如下：

	 {"errcode":0,"errmsg":"ok","chatdata":[{"seq":196,"msgid":"CAQQ2fbb4QUY0On2rYSAgAMgip/yzgs=","publickey_ver":3,"encrypt_random_key":"ftJ+uz3n/z1DsxlkwxNgE+mL38H42/KCvN8T60gbbtPD+Rta1hKTuQPzUzO6Hzne97MgKs7FfdDxDck/v8cDT6gUVjA2tZ/M7euSD0L66opJ/IUeBtpAtvgVSD5qhlaQjvfKJc/zPMGNK2xCLFYqwmQBZXbNT7uA69Fflm512nZKW/piK2RKdYJhRyvQnA1ISxK097sp9WlEgDg250fM5tgwMjujdzr7ehK6gtVBUFldNSJS7ndtIf6aSBfaLktZgwHZ57ONewWq8GJe7WwQf1hwcDbCh7YMG8nsweEwhDfUz+u8rz9an+0lgrYMZFRHnmzjgmLwrR7B/32Qxqd79A==","encrypt_chat_msg":"898WSfGMnIeytTsea7Rc0WsOocs0bIAerF6de0v2cFwqo9uOxrW9wYe5rCjCHHH5bDrNvLxBE/xOoFfcwOTYX0HQxTJaH0ES9OHDZ61p8gcbfGdJKnq2UU4tAEgGb8H+Q9n8syRXIjaI3KuVCqGIi4QGHFmxWenPFfjF/vRuPd0EpzUNwmqfUxLBWLpGhv+dLnqiEOBW41Zdc0OO0St6E+JeIeHlRZAR+E13Isv9eS09xNbF0qQXWIyNUi+ucLr5VuZnPGXBrSfvwX8f0QebTwpy1tT2zvQiMM2MBugKH6NuMzzuvEsXeD+6+3VRqL"}]}

	 *
	 * @return 返回是否调用成功
	 *      0   - 成功
	 *      !=0 - 失败	
	 */		
	 int GetChatData(WeWorkFinanceSdk_t* sdk, unsigned long long seq, unsigned int limit, const char *proxy,const char* passwd, int timeout,Slice_t* chatDatas);
"""
sdk_dll.GetChatData.argtypes = [ctypes.c_void_p, c_ulonglong, c_int, c_char_p, c_char_p, c_int, POINTER(Slice)]
sdk_dll.GetChatData.restype = c_int

"""
	/**
     * @brief 解析密文.企业微信自有解密内容
     * @param [in]  encrypt_key, getchatdata返回的encrypt_random_key,使用企业自持对应版本秘钥RSA解密后的内容
     * @param [in]  encrypt_msg, getchatdata返回的encrypt_chat_msg
     * @param [out] msg, 解密的消息明文
	 * @return 返回是否调用成功
	 *      0   - 成功
	 *      !=0 - 失败
     */
	 int DecryptData(const char* encrypt_key, const char* encrypt_msg, Slice_t* msg);
"""
sdk_dll.DecryptData.argtypes = [c_char_p, c_char_p, POINTER(Slice)]
sdk_dll.DecryptData.restype = c_int

"""
	/**
	 * 拉取媒体消息函数
	 * Return值=0表示该API调用成功
	 * 
	 *
	 * @param [in]  sdk				NewSdk返回的sdk指针
	 * @param [in]  sdkFileid		从GetChatData返回的聊天消息中，媒体消息包括的sdkfileid
	 * @param [in]  proxy			使用代理的请求，需要传入代理的链接。如：socks5://10.0.0.1:8081 或者 http://10.0.0.1:8081
	 * @param [in]  passwd			代理账号密码，需要传入代理的账号密码。如 user_name:passwd_123
	 * @param [in]  indexbuf		媒体消息分片拉取，需要填入每次拉取的索引信息。首次不需要填写，默认拉取512k，后续每次调用只需要将上次调用返回的outindexbuf填入即可。
	 * @param [in]  timeout			超时时间，单位秒
	 * @param [out] media_data		返回本次拉取的媒体数据.MediaData结构体.内容包括data(数据内容)/outindexbuf(下次索引)/is_finish(拉取完成标记)
	 
	 *
	 * @return 返回是否调用成功
	 *      0   - 成功
	 *      !=0 - 失败
	 */
	 int GetMediaData(WeWorkFinanceSdk_t* sdk, const char* indexbuf,
                     const char* sdkFileid,const char *proxy,const char* passwd, int timeout, MediaData_t* media_data);
"""
sdk_dll.GetMediaData.argtypes = [ctypes.c_void_p, c_char_p, c_char_p, c_char_p, c_char_p, c_int, POINTER(MediaData)]
sdk_dll.GetMediaData.restype = c_int

"""
    /**
     * @brief 释放sdk，和NewSdk成对使用
     * @return 
     */
	 void DestroySdk(WeWorkFinanceSdk_t* sdk);
"""
sdk_dll.DestroySdk.argtypes = [ctypes.c_void_p]

"""
    //--------------下面接口为了其他语言例如python等调用c接口，酌情使用--------------
    Slice_t* NewSlice();
"""
sdk_dll.NewSlice.restype = POINTER(Slice)
"""
    /**
     * @brief 释放slice，和NewSlice成对使用
     * @return 
     */
	 void FreeSlice(Slice_t* slice);
"""
sdk_dll.FreeSlice.argtypes = [POINTER(Slice)]

"""
    /**
     * @brief 为其他语言提供读取接口
     * @return 返回buf指针
     *     !=NULL - 成功
     *     NULL   - 失败
     */
	 char* GetContentFromSlice(Slice_t* slice);
"""
sdk_dll.GetContentFromSlice.argtypes = [POINTER(Slice)]
sdk_dll.GetContentFromSlice.restype = ctypes.c_char_p

"""
	 int GetSliceLen(Slice_t* slice);
"""
sdk_dll.GetSliceLen.argtypes = [POINTER(Slice)]
sdk_dll.GetSliceLen.restype = ctypes.c_int


"""
     // 媒体记录工具
     MediaData_t*  NewMediaData();
     void FreeMediaData(MediaData_t* media_data);
	 char* GetOutIndexBuf(MediaData_t* media_data);
	 char* GetData(MediaData_t* media_data);
	 int GetIndexLen(MediaData_t* media_data);
	 int GetDataLen(MediaData_t* media_data);
	 int IsMediaDataFinish(MediaData_t* media_data);
"""
sdk_dll.NewMediaData.restype = POINTER(MediaData)
sdk_dll.FreeMediaData.argtypes = [POINTER(MediaData)]
sdk_dll.GetOutIndexBuf.argtypes = [POINTER(MediaData)]
sdk_dll.GetOutIndexBuf.restype = ctypes.c_char_p
sdk_dll.GetData.argtypes = [POINTER(MediaData)]
sdk_dll.GetData.restype = ctypes.c_void_p
sdk_dll.GetIndexLen.argtypes = [POINTER(MediaData)]
sdk_dll.GetIndexLen.restype = ctypes.c_int
sdk_dll.GetDataLen.argtypes = [POINTER(MediaData)]
sdk_dll.GetDataLen.restype = ctypes.c_int
sdk_dll.IsMediaDataFinish.argtypes = [POINTER(MediaData)]
sdk_dll.IsMediaDataFinish.restype = ctypes.c_int

class WeWorkFinanceSdk:
    def __init__(self, corp_id_, corp_key_):
        self.sdk = sdk_dll.NewSdk()
        ret = sdk_dll.Init(self.sdk, corp_id_.encode(), corp_key_.encode())
        if ret != 0:
            raise Exception(f"Init SDK failed with error code: {ret}")


    def get_chat_data(self, seq, limit, proxy="", passwd="", timeout=30):
        """
        获取聊天数据
        :param seq:
        :param limit:
        :param proxy:
        :param passwd:
        :param timeout:
        :return:
        """
        chat_datas = sdk_dll.NewSlice()
        ret = sdk_dll.GetChatData(self.sdk, seq, limit, proxy.encode(), passwd.encode(), timeout, chat_datas)
        if ret != 0:
            print(f"Get data failed, return code: {ret}.")
            sdk_dll.FreeSlice(chat_datas)
            raise Exception(f"Get data failed with error code: {ret}")
        data = chat_datas.contents.buf[:chat_datas.contents.len]
        length_ = chat_datas.contents.len
        sdk_dll.FreeSlice(chat_datas)
        return data, length_


    def pull_media_file(self, file_id:str, proxy="", passwd="", timeout=30, max_retries=3):
        index_buf = create_string_buffer(512 * 1024)
        total_data = bytearray()
        is_finish = 0
        retries = 0
        while not is_finish and retries < max_retries:
            media_data = sdk_dll.NewMediaData()
            ret = sdk_dll.GetMediaData(self.sdk, index_buf.raw, file_id.encode(), proxy.encode(), passwd.encode(), timeout, media_data)
            if ret != 0:
                print(f"PullMediaData err ret: {ret}, retrying ({retries + 1}/{max_retries})...")
                retries += 1
                sdk_dll.FreeMediaData(media_data)
                time.sleep(3)
                continue
            # 获取二进制数据
            data = ctypes.string_at(media_data.contents.data, media_data.contents.data_len)
            # 存放到内存中
            total_data.extend(data)
            # 获取下一次调用的index_buf
            index_buf.raw = media_data.contents.outindexbuf[:media_data.contents.out_len]
            # 获取finish标记
            is_finish = media_data.contents.is_finish
            # 释放内存
            sdk_dll.FreeMediaData(media_data)
        return bytes(total_data), len(total_data)


    def download_media_file(self, file_id:str, file_save_path:str, md5sum="", proxy="", passwd="", timeout=30, max_retries=3):
        # 媒体文件每次拉取的最大size为512k，因此超过512k的文件需要分片拉取。若该文件未拉取完整，mediaData中的is_finish会返回0，同时mediaData中的outindexbuf会返回下次拉取需要传入GetMediaData的indexbuf。
        # indexbuf一般格式如右侧所示，”Range:bytes=524288-1048575“，表示这次拉取的是从524288到1048575的分片。单个文件首次拉取填写的indexbuf为空字符串，拉取后续分片时直接填入上次返回的indexbuf即可。
        index_buf, is_finish, retries = create_string_buffer(512 * 1024), 0, 0
        file_save_path_tmp = f'{file_save_path}.wxtmp'
        if len(md5sum) > 0:
            hmd5 = hashlib.md5()

        while not is_finish and retries < max_retries:
            media_data = sdk_dll.NewMediaData()
            ret = sdk_dll.GetMediaData(self.sdk, index_buf.raw, file_id.encode(), proxy.encode(), passwd.encode(), timeout, media_data)
            if ret != 0:
                print(f"PullMediaData err ret: {ret}, retrying ({retries + 1}/{max_retries})...")
                retries += 1
                # 单个分片拉取失败建议重试拉取该分片，避免从头开始拉取。
                sdk_dll.FreeMediaData(media_data)
                time.sleep(3)
                continue
             # 二进制数据写入文件
            with open(file_save_path_tmp, 'ab') as dstf:
                data = ctypes.string_at(media_data.contents.data, media_data.contents.data_len)
                dstf.write(ctypes.string_at(media_data.contents.data, media_data.contents.data_len))
                if len(md5sum) > 0:
                    hmd5.update(data)
            
            # 获取下一次调用的index_buf
            index_buf.raw = media_data.contents.outindexbuf[:media_data.contents.out_len]
            # 获取finish标记
            is_finish = media_data.contents.is_finish
            # 释放内存
            sdk_dll.FreeMediaData(media_data)

        md5_check_success = True
        if len(md5sum) > 0:
            if md5sum != hmd5.hexdigest():
                md5_check_success = False
        
        download_success = is_finish and retries < max_retries and md5_check_success
        if not download_success:
            # 下载失败，删除临时文件
            os.remove(file_save_path_tmp)
        else:
            # 下载成功, 修改临时文件名
            os.rename(file_save_path_tmp, file_save_path)
        return download_success

    @staticmethod
    def decrypt_data(encrypt_key, encrypt_chat_msg):
        """
        解密数据
        :param encrypt_key:
        :param encrypt_chat_msg:
        :return:
        """
        msgs = sdk_dll.NewSlice()
        ret = sdk_dll.DecryptData(encrypt_key.encode(), encrypt_chat_msg.encode(), msgs)
        if ret != 0:
            print("Decrypt data failed.")
            sdk_dll.FreeSlice(msgs)
            raise Exception(f"Decrypt data failed with error code: {ret}")

        data = msgs.contents.buf[:msgs.contents.len]
        length_ = msgs.contents.len
        sdk_dll.FreeSlice(msgs)
        return data, length_

    def destroy_sdk(self):
        sdk_dll.DestroySdk(self.sdk)


# 示例调用
if __name__ == "__main__":
    # 假设corpid和key是有效的
    corp_id = "your_corpid"
    corp_key = "your_key"
    start_seq = 0
    limit = 10
    has_prikey = False
    try:
        sdk = WeWorkFinanceSdk(corp_id, corp_key)
        # 导入私钥
        if os.path.exists(const_prikey_pem_path):
            with open(const_prikey_pem_path) as pk_file:
                privatekey = pk_file.read()
            # 初始化RSA
            rsakey = RSA.importKey(privatekey)
            cipher = PKCS1_v1_5.new(rsakey)
            has_prikey = True


        while True:
            # 获取聊天数据
            chat_data, length = sdk.get_chat_data(seq=start_seq, limit=limit)
            if chat_data is None:
                raise Exception(f"调用接口失败")

            ret_data = json.loads(chat_data)
            if ret_data.get("errcode") != 0:
                raise Exception(f"调用接口失败:{ret_data}")

            origin_data_list = ret_data.get("chatdata")
            if len(origin_data_list) <= 0:
                raise Exception(f"会话存档数据为空")

            # 获取最大的seq
            start_seq = max([p.get('seq') for p in origin_data_list])

            for chat_data in origin_data_list:
                if not has_prikey:
                    print(f'密文聊天数据: {chat_data}')
                    continue

                rdkey_str = chat_data.get("encrypt_random_key")
                rdkey_decoded = base64.b64decode(rdkey_str)

                encrypt_key = str(bytes.decode(cipher.decrypt(rdkey_decoded, None)))
                print(f'解密后的random_key: {encrypt_key}')
                encrypt_msg = chat_data.get("encrypt_chat_msg")
                byte_details, length = sdk.decrypt_data(encrypt_key, encrypt_msg)
                data_details = json.loads(byte_details)
                print(f'解密后的数据: {data_details}')

                if data_details.get('msgtype') is None:
                    continue
                elif data_details.get('msgtype') == 'text':
                    print(f'Text: {data_details.get("text").get("content")}')
                elif data_details.get('msgtype') == 'file':
                    """
                    "file": {
                        "md5sum": "72eae1fc232f042b5522ffb5a09db51d",
                        "filename": "测试文件.txt",
                        "fileext": "txt",
                        "filesize": 281,
                        "sdkfileid": "lYWExM4MzY="
                    }
                    """
                    file_info = data_details.get("file")
                    fileid = file_info.get('sdkfileid')
                    filename = file_info.get('filename')
                    filelen = file_info.get('filesize')
                    md5sum = file_info.get('md5sum')
                    file_content, length = sdk.pull_media_file(file_id=fileid)
                    if len(file_content) == filelen:
                        with open(f'./{filename}', 'wb') as dstf:
                            dstf.write(file_content)
                    else:
                        raise Exception(f"文件下载失败")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if 'sdk' in locals():
            sdk.destroy_sdk()