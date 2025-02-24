import requests
import json
import argparse
import random
import time
from curl_cffi import requests as cffi_requests
import sys
import os
import pandas as pd

# 类型映射
TYPE_MAPPING = {"web": 1, "app": 6, "miniapp": 7}

def get_current_time_filename():
    """生成当前时间戳的文件名"""
    return f"results_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"

def get_custom_headers():
    """获取用户自定义的HTTP请求头"""
    custom_headers = {}
    print("请输入 HTTP 请求头（若无需添加请直接回车）")
    custom_headers["Cookie"] = input("Cookie: ") or None
    custom_headers["Sign"] = input("Sign: ") or None
    custom_headers["Uuid"] = input("Uuid: ") or None
    custom_headers["Token"] = input("Token: ") or None
    return {k: v for k, v in custom_headers.items() if v is not None}

def generate_modern_headers():
    """生成现代浏览器的请求头"""
    browser_version = random.choice(["124", "123", "122"])
    platform = random.choice(["Windows", "macOS"])
    sec_platform = f"\"{platform}\""
    return {
        "Host": "hlwicpfwc.miit.gov.cn",
        "Sec-Ch-Ua": f"\"Chromium\";v=\"{browser_version}\", \"Google Chrome\";v=\"{browser_version}\", \"Not-A.Brand\";v=\"99\"",
        "Sec-Ch-Ua-Mobile": "?0",
        "Sec-Ch-Ua-Platform": sec_platform,
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": f"Mozilla/5.0 ({'Windows NT 10.0; Win64; x64' if 'Windows' in platform else 'Macintosh; Intel Mac OS X 10_15_7'}) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{browser_version}.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7",
        "Priority": "u=1, i",
        "Referer": "https://beian.miit.gov.cn/",
        "Origin": "https://beian.miit.gov.cn",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-site",
        "Connection": "keep-alive"
    }

def send_post_request(unit_name, headers, service_type=1, proxy=None):
    """发送POST请求"""
    url = "https://hlwicpfwc.miit.gov.cn/icpproject_query/api/icpAbbreviateInfo/queryByCondition"
    base_headers = generate_modern_headers()
    base_headers.update(headers)
    
    payload = {"pageNum": "", "pageSize": "", "unitName": unit_name, "serviceType": service_type}

    try:
        proxies = {"http": proxy, "https": proxy} if proxy else None
        response = cffi_requests.post(
            url,
            headers=base_headers,
            json=payload,
            impersonate="chrome110",
            proxies=proxies
        )
        if "Set-Cookie" in response.headers:
            base_headers["Cookie"] = response.headers["Set-Cookie"]
        return response
    except Exception as e:
        print(f"请求异常: {str(e)}")
        return None

def process_response(response_data, service_type):
    """处理响应数据"""
    results = []
    if response_data.get("success"):
        for item in response_data["params"]["list"]:
            result = {
                "unitName": item.get("unitName"),
                "mainLicence": item.get("mainLicence"),
                "serviceLicence": item.get("serviceLicence"),
                "updateRecordTime": item.get("updateRecordTime")
            }
            if service_type == 1:
                result["domain"] = item.get("domain")
            else:
                result.update({
                    "serviceName": item.get("serviceName"),
                    "leaderName": item.get("leaderName"),
                    "mainUnitAddress": item.get("mainUnitAddress")
                })
            results.append(result)
    return results

def write_to_excel(results_dict, output_file=None):
    """将结果写入Excel文件"""
    if not output_file:
        output_file = get_current_time_filename()
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, data in results_dict.items():
            if data:
                df = pd.DataFrame(data)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"结果已保存至：{output_file}")

def load_proxies():
    """从文件加载代理列表"""
    if not os.path.exists("proxy.txt"):
        return []
    with open("proxy.txt", "r") as f:
        return [line.strip() for line in f if line.strip()]

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='ICP备案查询工具')
    parser.add_argument('unit_name', nargs='?', help='查询的单位名称')
    parser.add_argument('-f', '--file', help='从文件读取单位名称列表')
    parser.add_argument('-o', '--output', help='输出文件名')
    parser.add_argument('-t', '--type', choices=['web', 'app', 'miniapp', 'all'], default='web', help='查询类型')
    parser.add_argument('-p', '--proxy_rotate', type=int, help='每N个请求更换代理')
    args = parser.parse_args()
    custom_headers = get_custom_headers()

    available_proxies = load_proxies()
    use_proxy = args.proxy_rotate is not None  # 是否启用代理

    # 代理轮换参数检查
    if use_proxy and not available_proxies:
        print("错误：代理轮换已启用，但未找到有效代理。请检查proxy.txt文件。")
        sys.exit(1)

    proxy_index = 0
    request_count = 0
    last_proxy_used = None
    query_types = ["web", "app", "miniapp"] if args.type == "all" else [args.type]
    units = []
    if args.file:
        with open(args.file, 'r', encoding='utf-8') as f:
            units = [line.strip() for line in f if line.strip()]
    elif args.unit_name:
        units = [args.unit_name]

    all_results = {t: [] for t in query_types}
    blocked = False

    try:
        for unit_idx, unit in enumerate(units):
            print(f"\n正在查询第 {unit_idx+1}/{len(units)} 个单位：{unit}")
            
            for type_idx, query_type in enumerate(query_types):
                service_type = TYPE_MAPPING[query_type]
                print(f"正在查询 {query_type} 类型...")
                max_retries = 5 if use_proxy else 0  # 未启用代理时不重试
                retry_count = 0
                success = False
                current_proxy = None

                while retry_count <= max_retries and not success:
                    # 确定当前代理
                    if use_proxy and available_proxies:
                        if args.proxy_rotate and (request_count % args.proxy_rotate == 0):
                            proxy_index = (proxy_index + 1) % len(available_proxies)
                        current_proxy = available_proxies[proxy_index]
                    else:
                        current_proxy = None

                    # 打印代理变更信息
                    if current_proxy != last_proxy_used and current_proxy is not None:
                        print(f"使用代理：{current_proxy}")
                        last_proxy_used = current_proxy

                    # 发送请求
                    response = send_post_request(unit, custom_headers, service_type, current_proxy)
                    request_count += 1

                    # 403检测逻辑（仅在未使用代理时启用）
                    if response and response.status_code == 403:
                        if not current_proxy:  # 未使用代理时处理
                            print("\n⚠️ 访问被拒绝（403），可能触发防护机制")
                            write_to_excel(all_results, args.output)
                            sys.exit(1)
                        else:  # 使用代理时的处理
                            print(f"⚠️ 代理 {current_proxy} 返回403，继续尝试其他代理...")
                            if current_proxy in available_proxies:
                                available_proxies.remove(current_proxy)
                            if available_proxies:
                                proxy_index = proxy_index % len(available_proxies)
                            else:
                                print("所有代理均已失效，正在保存数据并退出...")
                                write_to_excel(all_results, args.output)
                                sys.exit(1)
                            continue  # 跳过后续处理，直接重试

                    if response and response.status_code == 200:
                        try:
                            response_data = response.json()
                            if response_data.get("code") == 401:
                                print("Token已过期，访问被拒绝。")
                                choice = input("请选择操作：\n1. 更新请求头,继续任务 \n2. 中断并保存结果\n请输入选项：")
                                if choice == '1':
                                    new_headers = get_custom_headers()
                                    custom_headers.update(new_headers)
                                    print("请求头已更新，继续查询...")
                                else:
                                    print("用户选择中断，正在保存数据...")
                                    blocked = True
                                    break
                            results = process_response(response_data, service_type)
                            all_results[query_type].extend(results)
                            success = True
                            # 调整延迟时间
                            delay = random.uniform(0.5, 1.5) if not use_proxy else random.uniform(0.2, 0.5)
                            time.sleep(delay)
                            print(f"随机延迟: {delay:.2f}秒")
                        except Exception as e:
                            print(f"响应解析失败: {str(e)}")
                    else:
                        if current_proxy and use_proxy:
                            if current_proxy in available_proxies:
                                available_proxies.remove(current_proxy)
                                print(f"代理 {current_proxy} 失效，已移除")
                                if available_proxies:
                                    proxy_index = proxy_index % len(available_proxies)
                                else:
                                    print("无可用代理，正在保存数据并退出...")
                                    write_to_excel(all_results, args.output)
                                    sys.exit(1)
                        # 未启用代理或代理用尽时直接退出循环
                        if not use_proxy:
                            print("请求失败（未启用代理，不进行重试）")
                            break
                    retry_count += 1

                if not success:
                    print(f"{query_type} 类型查询失败")
                if blocked:
                    break
            if blocked:
                break
        
    except KeyboardInterrupt:
        print("\n用户中断操作，正在保存数据...")
        write_to_excel(all_results, args.output)
    
    write_to_excel(all_results, args.output)

if __name__ == '__main__':
    main()