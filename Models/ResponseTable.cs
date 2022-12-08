namespace ObjectDataMapper.Models;

/// <summary>
/// 테이블 표현 모델 
/// </summary>
public class ResponseTable
{
    /// <summary>
    /// 테이블명
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 컬럼 정보 
    /// </summary>
    public List<ResponseColumn> columns { get; set; } = new List<ResponseColumn>();

    /// <summary>
    /// 모든 컬럼이 각자 컬럼을 보유하고 있는지
    /// </summary>
    public bool HasCommentInAllColumns { get; set; }
}

/// <summary>
/// 컬럼 표현 모델
/// </summary>
public class ResponseColumn
{
    /// <summary>
    /// 컬럼명
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 주석내용
    /// </summary>
    public string Comments { get; set; }
}