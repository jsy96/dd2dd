'use client';

import { useState, useCallback } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Upload, Download, FileSpreadsheet, FileText, Loader2, CheckCircle2, AlertCircle } from 'lucide-react';
import { toast } from 'sonner';

interface ProcessingResult {
  success: boolean;
  message: string;
  wordFileUrl?: string;
  excelFileUrl?: string;
}

export default function Home() {
  const [manifestFile, setManifestFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [result, setResult] = useState<ProcessingResult | null>(null);

  const handleFileChange = useCallback((file: File | null) => {
    setManifestFile(file);
    setResult(null);
  }, []);

  const handleProcess = async () => {
    if (!manifestFile) {
      toast.error('请上传舱单文件');
      return;
    }

    setIsProcessing(true);
    setResult(null);

    try {
      const formData = new FormData();
      formData.append('manifest', manifestFile);

      const response = await fetch('/api/process', {
        method: 'POST',
        body: formData,
      });

      const data: ProcessingResult = await response.json();
      setResult(data);

      if (data.success) {
        toast.success('文件处理成功');
      } else {
        toast.error(data.message || '处理失败');
      }
    } catch (error) {
      console.error('处理失败:', error);
      toast.error('处理失败，请重试');
      setResult({
        success: false,
        message: '处理失败，请重试',
      });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownload = async (url: string, filename: string) => {
    try {
      const response = await fetch(url);
      const blob = await response.blob();
      const blobUrl = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = blobUrl;
      link.download = filename;
      link.click();
      window.URL.revokeObjectURL(blobUrl);
    } catch (error) {
      console.error('下载失败:', error);
      toast.error('下载失败');
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-purple-50 dark:from-gray-900 dark:via-gray-800 dark:to-gray-900 flex items-center justify-center p-4">
      <div className="w-full max-w-2xl">
        {/* 标题 */}
        <div className="text-center mb-8">
          <h1 className="text-3xl sm:text-4xl font-bold text-gray-900 dark:text-white mb-2">
            舱单文件处理系统
          </h1>
          <p className="text-gray-600 dark:text-gray-300">
            上传舱单文件，自动生成提单确认件和装箱单发票
          </p>
        </div>

        {/* 主卡片 */}
        <Card className="shadow-xl">
          <CardHeader className="text-center">
            <CardTitle className="text-xl flex items-center justify-center gap-2">
              <FileSpreadsheet className="w-5 h-5 text-blue-600" />
              上传舱单文件
            </CardTitle>
            <CardDescription>
              支持 .xls 或 .xlsx 格式的舱单文件
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-6">
            {/* 文件上传 */}
            <div className="space-y-2">
              <Input
                type="file"
                accept=".xls,.xlsx"
                onChange={(e) => handleFileChange(e.target.files?.[0] || null)}
                className="cursor-pointer"
              />
              {manifestFile && (
                <div className="flex items-center gap-2 text-sm text-green-600 dark:text-green-400 justify-center">
                  <CheckCircle2 className="w-4 h-4" />
                  {manifestFile.name}
                </div>
              )}
            </div>

            {/* 处理按钮 */}
            <Button
              onClick={handleProcess}
              disabled={!manifestFile || isProcessing}
              className="w-full h-12 text-base"
              size="lg"
            >
              {isProcessing ? (
                <>
                  <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                  处理中...
                </>
              ) : (
                <>
                  <Upload className="w-4 h-4 mr-2" />
                  开始处理
                </>
              )}
            </Button>

            {/* 处理失败提示 */}
            {result && !result.success && (
              <div className="flex flex-col items-center justify-center py-4 text-red-500">
                <AlertCircle className="w-8 h-8 mb-2" />
                <p>{result.message}</p>
              </div>
            )}

            {/* 下载区域 */}
            {result?.success && result.wordFileUrl && result.excelFileUrl && (
              <div className="pt-4 border-t">
                <p className="text-center text-sm text-gray-500 dark:text-gray-400 mb-4">
                  文件已生成，点击下载
                </p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  {/* 提单确认件 */}
                  <Card className="border-2 border-purple-200 dark:border-purple-800 hover:shadow-lg transition-shadow">
                    <CardContent className="pt-6">
                      <div className="flex flex-col items-center text-center">
                        <FileText className="w-12 h-12 text-purple-600 mb-3" />
                        <h3 className="font-semibold mb-3">提单确认件</h3>
                        <Button
                          onClick={() => handleDownload(result.wordFileUrl!, '提单确认件.doc')}
                          className="w-full"
                          variant="default"
                        >
                          <Download className="w-4 h-4 mr-2" />
                          下载
                        </Button>
                      </div>
                    </CardContent>
                  </Card>

                  {/* 装箱单发票 */}
                  <Card className="border-2 border-green-200 dark:border-green-800 hover:shadow-lg transition-shadow">
                    <CardContent className="pt-6">
                      <div className="flex flex-col items-center text-center">
                        <FileSpreadsheet className="w-12 h-12 text-green-600 mb-3" />
                        <h3 className="font-semibold mb-3">装箱单发票</h3>
                        <Button
                          onClick={() => handleDownload(result.excelFileUrl!, '装箱单发票.xls')}
                          className="w-full"
                          variant="default"
                        >
                          <Download className="w-4 h-4 mr-2" />
                          下载
                        </Button>
                      </div>
                    </CardContent>
                  </Card>
                </div>
              </div>
            )}
          </CardContent>
        </Card>

        {/* 使用说明 */}
        <div className="mt-6 text-center text-sm text-gray-500 dark:text-gray-400">
          <p>使用流程：上传舱单 → 点击处理 → 下载生成的文件</p>
        </div>
      </div>
    </div>
  );
}
