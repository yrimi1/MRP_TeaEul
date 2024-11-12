
-- 테이블 변경  ( sb_Stock )

-- 필드추가
-- Stuffin.SubulWidthID, StuffAssign, OutWare, Order
--
--





SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO
















/********************************************************
 * Date : 2003-10-30 (THU)
 *
 * Description: 수주입력 수정 - 이장훈
 *		MRPPlus (COrder -> frmOrder)
 ********************************************************/
ALTER         PROCEDURE xp_Order_uOrder
	@OrderID        AS CHAR(10) OUTPUT,
	@CustomID	AS CHAR(4),
	@OrderNO	AS VARCHAR(24),
	@PoNo		AS VARCHAR(24),
	@OrderForm	AS CHAR(1),
	@OrderClss	AS CHAR(1),
	@AcptDate	AS CHAR(8),
	@DvlyDate	AS CHAR(8),
	@ArticleID	AS CHAR(4),
	@DvlyPlace	AS VARCHAR(40),
	@WorkID		AS CHAR(4),
	@PriceClss	AS CHAR(1),
	@ExchRate	AS NUMERIC(6,2),
	@OrderQty	AS NUMERIC(12),
	@UnitClss	AS CHAR(1),
	@ColorCnt	AS INTEGER,
	@StuffWidth	AS CHAR(2),
	@StuffWeight	AS INTEGER,
	@CutQty		AS INTEGER,
	@WorkWidth	AS CHAR(2),
	@WorkWeight	AS INTEGER,
	@WorkDensity	AS SMALLINT,
	@ChunkRate	AS NUMERIC(7,2),
	@LossRate	AS NUMERIC(7,2),
	@ReduceRate	AS NUMERIC(7,2),
	@TagClss	AS CHAR(1),
	@LabelID	AS CHAR(2),
	@BandID		AS CHAR(2),
	@EndClss	AS CHAR(1),
	@MadeClss	AS CHAR(1),
	@SurfaceClss	AS CHAR(1),
	@ShipClss	AS VARCHAR(10),
	@AdvnClss	AS VARCHAR(10),
	@LotClss	AS VARCHAR(10),
	@EndMark	AS VARCHAR(250),
	@TagArticle	AS VARCHAR(16),
	@TagOrderNo	AS VARCHAR(16),
	@TagRemark	AS VARCHAR(16),
	@TagRemark2	AS VARCHAR(16),
	@Tag		AS VARCHAR(250),
	@BasisID	AS CHAR(2),
	@BasisUnit	AS CHAR(1),
	@SpendingClss	AS CHAR(1),
	@DyeingID	AS CHAR(1),
	@WorkingClss	AS CHAR(1),
	@AccountClss	As CHAR(1),
	@Remark		AS VARCHAR(250),
	@ActiveClss	AS CHAR(1),
	@BTID		AS CHAR(8),
	@BTIDSeq	AS SMALLINT,
	@ChemClss	AS CHAR(1),
	@OrderFlag	AS CHAR(1),
        @PatternID      AS CHAR(2) = '',
        @Item           AS VARCHAR(50)='',
        @SubulWidthID   AS CHAR(2)
AS
	UPDATE [Order] SET
	CustomID = @CustomID, OrderNo = @OrderNo, PoNo = dbo.fn_CheckData(@PoNo), OrderForm = @OrderForm, OrderClss = @OrderClss, 
	AcptDate = @AcptDate, DvlyDate = @DvlyDate, ArticleID = @ArticleID, DvlyPlace = dbo.fn_CheckData(@DvlyPlace), 
	WorkID = @WorkID, PriceClss = @PriceClss, ExchRate = @ExchRate, OrderQty = @OrderQty, 
	UnitClss = @UnitClss, ColorCnt = @ColorCnt, StuffWidth = @StuffWidth, StuffWeight = @StuffWeight, CutQty = @CutQty,
	WorkWidth = @WorkWidth, WorkWeight = @WorkWeight, WorkDensity = @WorkDensity, ChunkRate = @ChunkRate, LossRate = @LossRate, 
	ReduceRate = @ReduceRate, TagClss = @TagClss, LabelID = @LabelID, BandID = @BandID, EndClss = @EndClss, MadeClss = @MadeClss, 
	SurfaceClss = @SurfaceClss, ShipClss = dbo.fn_CheckData(@ShipClss), AdvnClss = dbo.fn_CheckData(@AdvnClss), 
	LotClss = dbo.fn_CheckData(@LotClss), EndMark = dbo.fn_CheckData(@EndMark),
	TagArticle = dbo.fn_CheckData(@TagArticle), TagOrderNo = dbo.fn_CheckData(@TagOrderNo), 
	TagRemark = dbo.fn_CheckData(@TagRemark), TagRemark2 = dbo.fn_CheckData(@TagRemark2),
	Tag = dbo.fn_CheckData(@Tag),	BasisID = @BasisID, BasisUnit = @BasisUnit, 
	SpendingClss = @SpendingClss, DyeingID = @DyeingID, WorkingClss = @WorkingClss, AccountClss = @AccountClss,
	Remark = dbo.fn_CheckData(@Remark), ActiveClss = @ActiveClss, BTID = dbo.fn_CheckData(@BTID), BTIDSeq = @BTIDSeq,
	ChemClss = @ChemClss, OrderFlag = @OrderFlag, PatternID = @PatternID, Item = @Item, SubulWidthID = @SubulWidthID

	WHERE OrderID = @OrderID

	UPDATE [wk_Result] SET CustomID = @CustomID, ArticleID = @ArticleID WHERE OrderID = @OrderID

        DELETE pl_InputDET
          FROM ( SELECT InstDate, InstSeq
                   FROM Pl_Input
                  WHERE OrderID = @OrderID ) AS AA

         WHERE pl_InputDET.InstDate = AA.InstDate
           AND pl_InputDET.InstSeq = AA.InstSeq

     --   DELETE pl_Input
     --    WHERE OrderID = @OrderID







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

















/********************************************************
 * Date : 2003-10-30 (THU)
 *
 * Description: 수주입력 - 이장훈
 *		MRPPlus (COrder -> frmOrder)
 ********************************************************/
ALTER          PROCEDURE xp_Order_iOrder
	@OrderID        AS CHAR(10) OUTPUT,
	@CustomID	AS CHAR(4),
	@OrderNO	AS VARCHAR(24),
	@PoNo		AS VARCHAR(24),
	@OrderForm	AS CHAR(1),
	@OrderClss	AS CHAR(1),
	@AcptDate	AS CHAR(8),
	@DvlyDate	AS CHAR(8),
	@ArticleID	AS CHAR(4),
	@DvlyPlace	AS VARCHAR(40),
	@WorkID		AS CHAR(4),
	@PriceClss	AS CHAR(1),
	@ExchRate	AS NUMERIC(6,2),
	@OrderQty	AS NUMERIC(12),
	@UnitClss	AS CHAR(1),
	@ColorCnt	AS INTEGER,
	@StuffWidth	AS CHAR(2),
	@StuffWeight	AS INTEGER,
	@CutQty		AS INTEGER,
	@WorkWidth	AS CHAR(2),
	@WorkWeight	AS INTEGER,
	@WorkDensity	AS SMALLINT,
	@ChunkRate	AS NUMERIC(7,2),
	@LossRate	AS NUMERIC(7,2),
	@ReduceRate	AS NUMERIC(7,2),
	@TagClss	AS CHAR(1),
	@LabelID	AS CHAR(2),
	@BandID		AS CHAR(2),
	@EndClss	AS CHAR(1),
	@MadeClss	AS CHAR(1),
	@SurfaceClss	AS CHAR(1),
	@ShipClss	AS VARCHAR(10),
	@AdvnClss	AS VARCHAR(10),
	@LotClss	AS VARCHAR(10),
	@EndMark	AS VARCHAR(250),
	@TagArticle	AS VARCHAR(16),
	@TagOrderNo	AS VARCHAR(16),
	@TagRemark	AS VARCHAR(16),
	@TagRemark2	AS VARCHAR(16),
	@Tag		AS VARCHAR(250),
	@BasisID	AS CHAR(2),
	@BasisUnit	AS CHAR(1),
	@SpendingClss	AS CHAR(1),
	@DyeingID	AS CHAR(1),
	@WorkingClss	AS CHAR(1),
	@AccountClss	As CHAR(1),
	@Remark		AS VARCHAR(250),
	@ActiveClss	AS CHAR(1),
	@BTID		AS CHAR(8),
	@BTIDSeq	AS SMALLINT,
	@ChemClss	AS CHAR(1),
	@OrderFlag	AS CHAR(1),
        @PatternID      AS CHAR(2) ='',
        @Item           AS VARCHAR(50) ='',
        @SubulWidthID   AS CHAR(2) = ''

 

AS
	IF LEN(@OrderID) <> 10
	BEGIN
		SET @OrderID = dbo.fn_NewOrderID(Getdate())
	END

	INSERT INTO [Order] (OrderID, CustomID, OrderNo, PoNo, OrderForm, OrderClss, AcptDate, DvlyDate, ArticleID, 
		DvlyPlace, WorkID, PriceClss, ExchRate, OrderQty, UnitClss, ColorCnt,
		StuffWidth, StuffWeight, CutQty, WorkWidth, WorkWeight, WorkDensity, ChunkRate, LossRate, ReduceRate, 
		TagClss, LabelID, BandID, EndClss, MadeClss, SurfaceClss, ShipClss, 
		AdvnClss, LotClss, EndMark, TagArticle, 
		TagOrderNo, TagRemark, TagRemark2, Tag, BasisID, BasisUnit, 
		SpendingClss, DyeingID, WorkingClss, AccountClss, Remark, ActiveClss, CurrDate, 
		PatternID, ModifyClss, ModifyRemark, CancelRemark, CloseClss, CloseDate, StuffCloseClss, StuffCloseDate, ModifyDate,
		BTID, BTIDSeq, ChemClss, OrderFlag, Item, SubulWidthID )
	VALUES(@OrderID, @CustomID, @OrderNo, dbo.fn_CheckData(@PoNo), @OrderForm, @OrderClss, @AcptDate, @DvlyDate, @ArticleID,
		dbo.fn_CheckData(@DvlyPlace), @WorkID, @PriceClss, @ExchRate, @OrderQty, @UnitClss, @ColorCnt,
		@StuffWidth, @StuffWeight, @CutQty, @WorkWidth, @WorkWeight, @WorkDensity, @ChunkRate, @LossRate, @ReduceRate,
		@TagClss, @LabelID, @BandID, @EndClss, @MadeClss, @SurfaceClss, dbo.fn_CheckData(@ShipClss), 
		dbo.fn_CheckData(@AdvnClss), dbo.fn_CheckData(@LotClss), dbo.fn_CheckData(@EndMark), dbo.fn_CheckData(@TagArticle), 
		dbo.fn_CheckData(@TagOrderNo), dbo.fn_CheckData(@TagRemark), dbo.fn_CheckData(@TagRemark2), dbo.fn_CheckData(@Tag), @BasisID, @BasisUnit, 
		@SpendingClss, @DyeingID, @WorkingClss, @AccountClss, dbo.fn_CheckData(@Remark), @ActiveClss, dbo.fn_Date(Getdate()),
		@PatternID, '', '', '', '', '', '', '', '',
		dbo.fn_CheckData(@BTID), @BTIDSeq, @ChemClss, @OrderFlag, @Item, @SubulWidthID)

	-- 배색이 지정되지 않은 건들을 위해서 미확정 색상 추가
	INSERT INTO [OrderColor] (OrderID, OrderSeq, ColorID, Color, DesignNo, ColorQty, UnitPrice)
	VALUES(@OrderID, 0, '', '미확정', ' ', 0, 0)

        DELETE pl_InputDET
          FROM ( SELECT InstDate, InstSeq
                   FROM Pl_Input
                  WHERE OrderID = @OrderID ) AS AA

         WHERE pl_InputDET.InstDate = AA.InstDate
           AND pl_InputDET.InstSeq = AA.InstSeq

--        DELETE pl_Input
--         WHERE OrderID = @OrderID








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO











/********************************************************
 * Date : 2003-10-30 (THU)
 *
 * Description: 수주상세 - 이장훈
 *              (COrder -> frmOrder)
 ********************************************************/
ALTER    PROCEDURE xp_Order_sOrderOne
	@OrderID 	CHAR(10)
AS
	SELECT A.OrderID, A.CustomID, A.OrderNo, B.KCustom, A.PONO, A.OrderForm, A.OrderClss, A.Item,
		A.AcptDate, A.DvlyDate, A.ArticleID, C.Article, A.DvlyPlace, A.WorkID,
		A.PriceClss, A.ExchRate, A.OrderQty, A.UnitClss, A.ColorCnt, A.SubulWidthID, 
		A.StuffWidth, A.StuffWeight, A.CutQty, 	A.WorkWidth, A.WorkWeight, A.WorkDensity, 
		A.ChunkRate, A.LossRate, A.ReduceRate, A.TagClss, A.LabelID, A.BandID, A.EndClss, A.MadeClss,
		A.SurfaceClss, A.ShipClss, A.AdvnClss, A.LotClss, A.EndMark, A.TagArticle, A.TagOrderNo, 
		A.TagRemark, A.TagRemark2, A.Tag, A.BasisID, A.BasisUnit, A.SpendingClss, A.DyeingID, A.WorkingClss, A.PatternID, A.BTID, A.BTIDSeq, A.ChemClss,
		A.AccountClss, A.ModifyClss, A.ModifyRemark, A.CancelRemark, A.Remark, A.ActiveClss, A.CloseClss, A.ModifyDate, A.OrderFlag
	FROM [Order] A, [mt_Custom] B, [mt_Article] C
	WHERE A.CustomID = B.CustomID
		AND A.ArticleID = C.ArticleID
		AND A.OrderID = @OrderID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO










--*********************************************
CREATE           TRIGGER Order_UPD_Trigger
ON [Order]
FOR UPDATE
AS 
     -- 수주확정 테이블 넣기 ---  2004/04/10 10:19 수정 

      

      UPDATE OutWare
         SET SubulWidthID = BB.SubulWidthID
        FROM OutWare AA, ( SELECT OrderID, SubulWidthID 
                 FROM Inserted ) BB
       WHERE AA.OrderID = BB.OrderID


      UPDATE StuffIN
         SET SubulWidthID = BB.SubulWidthID
        FROM StuffIN AA, ( SELECT OrderID, SubulWidthID 
                 FROM Inserted ) BB
       WHERE AA.OrderID = BB.OrderID


      UPDATE StuffAssign
         SET SubulWidthID = BB.SubulWidthID
        FROM StuffIN AA, ( SELECT OrderID, SubulWidthID 
                 FROM Inserted ) BB
       WHERE AA.OrderID = BB.OrderID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







--**********************************************************
-- 재고 명세서
--
--
--
--
--**********************************************************
ALTER           PROCEDURE xp_sbStock_StockReport
             @sDate          CHAR(8) 
           , @nChkCustom     SMALLINT = 0
           , @CustomID       CHAR(4) = ''
           , @sTaxClss       CHAR(1)      -- '0':A건, '1':B건, '9':전체
AS
/*

DECLARE      @sDate          CHAR(8) 
           , @nChkCustom     SMALLINT 
           , @CustomID       CHAR(4)
   

SELECT @sDate     = '20040401' 
           , @nChkCustom   = 0
           , @CustomID     = ''

*/
      SELECT AA.kCustom, Depth='Z0', BB.Article
           , SubulWidth = ISNULL( ( SELECT StuffWidth
                                        FROM mt_StuffWidth ZZ
                                       WHERE CC.SubulWidthID = ZZ.StuffWidthID ), '' )
                   
           , ProcInClss = ISNULL( ( SELECT WorkName
                                       FROM mt_Work ZZ
                                      WHERE CC.ProcInClss = ZZ.WorkID ), '' )

           , StockQty = ISNULL( StockQty + StuffINQty - OutQty, 0 )
        FROM [mt_Custom] AA
           , [mt_Article] BB
           , fn_sbStock_sStock ( @sDate, @nChkCustom, @CustomID, 0, '', '', @sTaxClss ) AS  CC
       WHERE AA.CustomID = CC.CustomID
         AND CC.ArticleID = BB.ArticleID
         AND ISNULL( StockQty + StuffINQty - OutQty, 0 ) <> 0

      UNION ALL

      SELECT AA.kCustom, Depth='Z1', Article='거래처계'
           , SubulWidth = ''
           , ProcInClss='', StockQty = ISNULL ( SUM(StockQty + StuffINQty - OutQty) , 0 )
        FROM [mt_Custom] AA
           , [mt_Article] BB
           , fn_sbStock_sStock ( @sDate, @nChkCustom, @CustomID, 0, '', '', @sTaxClss ) AS  CC
       WHERE AA.CustomID = CC.CustomID
         AND CC.ArticleID = BB.ArticleID
         AND ISNULL( StockQty + StuffINQty - OutQty, 0 ) <> 0
    GROUP BY AA.kCustom

      UNION ALL



      SELECT kCustom='ZZZZZZZZ', Depth='Z2', Article= '업체수 : ' + CONVERT( CHAR(3), COUNT( DISTINCT CC.CUSTOMID ) ), SubulWidth=''
          , ProcInClss='', StockQty = ISNULL ( SUM(StockQty + StuffINQty - OutQty) , 0 ) 
        FROM [mt_Custom] AA
           , [mt_Article] BB
           , fn_sbStock_sStock ( @sDate, @nChkCustom, @CustomID, 0, '', '', @sTaxClss ) AS  CC
       WHERE AA.CustomID = CC.CustomID
         AND CC.ArticleID = BB.ArticleID
         AND ISNULL( StockQty + StuffINQty - OutQty, 0 ) <> 0

    ORDER BY AA.kCustom, Depth, BB.Article     













GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO








-- exec xp_Subul_iStock  '20040831', '0400', '0001', '0001', '1', 90, '1', '0'

/********************************************************
 * Date : 2004-03-24(WED)
 *
 * Description: 재고 입력
 *		MRPPlus2 (CSubul -> frmControlStock)
 ********************************************************/
ALTER         PROCEDURE xp_Subul_iStock
	@sDate		CHAR(8),
	@CustomID	CHAR(4),
	@ArticleID	CHAR(4),
	@ProcInClss	CHAR(4),
	@StockClss	CHAR(1), 
	@StockQty	NUMERIC(9,1), 
	@StockUnitClss  CHAR(1),
        @STaxClss       CHAR(1),
        @SubulWidthID   CHAR(2)
AS


	INSERT INTO [sb_Stock] (BasisDate, CustomID, ArticleID, ProcInClss, StockClss, StockQty, StockUnitClss, TaxClss, SubulWidthID)
	VALUES (@sDate, @CustomID, @ArticleID, @ProcInClss, @StockClss, @StockQty, @StockUnitClss, @STaxClss, @SubulWidthID )






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO









/********************************************************
 * Date : 2004-03-24(WED)
 *
 * Description: 재고 수정
 *		MRPPlus2 (CSubul -> frmControlStock)
 ********************************************************/
ALTER        PROCEDURE xp_Subul_uStock
	@sDate		CHAR(8),
	@CustomID	CHAR(4),
	@ArticleID	CHAR(4),
	@ProcInClss	CHAR(4),
	@StockClss	CHAR(1), 
	@StockQty	NUMERIC(12,0), 
	@StockUnitClss  CHAR(1),
        @STaxClss       CHAR(1),
        @SubulWidthID   CHAR(2)
AS

	UPDATE [sb_Stock] 
           SET ProcInClss = @ProcInClss
             , StockClss = @StockClss
             , StockQty = @StockQty
             , StockUnitClss = @StockUnitClss
	WHERE BasisDate = @sDate 
          AND CustomID = @CustomID 
          AND ArticleID = @ArticleID
          AND SubulWidthID = @SubulWidthID
          AND TaxClss = @STaxClss




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO









/********************************************************
 * Date : 2004-03-24(WED)
 *
 * Description: 재고 삭제
 *		MRPPlus2 (CSubul -> frmControlStock)
 ********************************************************/
ALTER        PROCEDURE xp_Subul_dStock
	@sDate		CHAR(8),
	@CustomID	CHAR(4),
	@ArticleID	CHAR(4),
        @TaxClss        CHAR(1),
        @SubulWidthID   CHAR(2)
AS

	DELETE FROM [sb_Stock] 
	      WHERE BasisDate = @sDate 
                AND CustomID = @CustomID 
                AND ArticleID = @ArticleID
                AND TaxClss = @TaxClss
                AND SubulWidthID = @SubulWidthID





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO









/********************************************************
 * Date : 2004-03-24(WED)
 *
 * Description: 재고 리스트
 *		MRPPlus2 (CSubul -> frmControlStock)
 ********************************************************/
ALTER        PROCEDURE xp_Subul_sStockDataOne
	@sDate		CHAR(8), 
	@CustomID	CHAR(4) = '',
	@ArticleID	CHAR(4) = '',
        @TaxClss        CHAR(1) = '',
        @SubulWidthID   CHAR(2) = ''
AS
	
	SELECT *
	  FROM [sb_Stock] 
         WHERE BasisDate = @sDate 
           AND CustomID = @CustomID 
           AND ArticleID = @ArticleID
           AND SubulWidthID = @SubulWidthID
           AND TaxClss = @TaxClss








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO




-- exec xp_Subul_sStockList  '20041231', '20041231'





/********************************************************
 * Date : 2004-03-24(WED)
 *
 * Description: 재고 리스트
 *		MRPPlus2 (CSubul -> frmControlStock)
 ********************************************************/
ALTER        PROCEDURE xp_Subul_sStockList
	@StartDate	CHAR(8), 
	@EndDate	CHAR(8),
	@ChkCustom	TINYINT = 0, 
	@CustomID	CHAR(4) = '',
	@ChkArticle	TINYINT = 0,
	@ArticleID	CHAR(4) = ''
AS
/*

DECLARE 	@StartDate	CHAR(8), 
	@EndDate	CHAR(8),
	@ChkCustom	TINYINT, 
	@CustomID	CHAR(4),
	@ChkArticle	TINYINT,
	@ArticleID	CHAR(4)

SELECT @StartDate	= '20041231' , 
	@EndDate	= '20041231',
	@ChkCustom	= 0,
	@CustomID	='',
	@ChkArticle	=0,
	@ArticleID	=''
*/

        SELECT A.*, B.KCustom, C.Article, D.StuffWidth
          FROM [sb_Stock] A, [mt_Custom] B, [mt_Article] C,  [mt_StuffWidth] D
         WHERE A.CustomID = B.CustomID AND A.ArticleID = C.ArticleID  AND A.SubulwidthID = D.StuffWidthID
           AND BasisDate BETWEEN @StartDate AND  @EndDate
           AND ( ( @ChkCustom > 0  AND  A.CustomID = @CustomID )
              OR ( @ChkCustom = 0  ) )
           AND ( ( @ChkArticle > 0  AND  A.ArticleID = @ArticleID )
              OR ( @ChkArticle = 0  ) )


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO









--********************************************************************
-- 수불명세서
--
-- fn_sbStock_sStock 사용
--
--********************************************************************
ALTER               PROCEDURE xp_Subul_sReport
       @sDate             CHAR(8)
     , @eDate             CHAR(8)
     , @nChkCustom        SMALLINT = 0
     , @CustomID          CHAR(4)  =''   
     , @nChkArticle       SMALLINT = 0
     , @ArticleID         CHAR(4)  = ''
     , @SubulWidthID      CHAR(2)  = ''       -- 재고폭
     , @TaxClss           CHAR(1)  = '9'      -- '0' :A /  '1' : b / '9': 전체

AS
/*
      FROM ( SELECT BB.kCustom, CC.Article, IODate = '', Cls = '0', Custom='', StuffRoll=0, StuffQty = ISNULL( StockQty +  StuffINQty -  OutQty, 0 )
                  , OrderNO='', OutRoll=0, OutQty=0, OutRealQty=0, Remark='', pkey=''
               FROM DBO.fn_sbStock_sStock(@sDate, @nChkCustom, @CustomID, @nChkArticle, @ArticleID )  AA
                  , [mt_Custom] BB, [mt_Article] CC
              WHERE AA.CustomID = BB.CustomID
                AND AA.ArticleID = CC.ArticleID   

              UNION ALL


exec xp_Subul_sReport  '20041201', '20041227', 1, '0052'
*/
/*
declare
              @sDate             CHAR(8)
            , @eDate             CHAR(8)
            , @nChkCustom        SMALLINT
            , @CustomID          CHAR(4)  
            , @nChkArticle       SMALLINT
            , @ArticleID         CHAR(4)  
            , @TaxClss           CHAR(1)        -- '0' :A /  '1' : b / '9': 전체

select @sDate      = '20040901'
            , @eDate   = '20040916'
            , @nChkCustom  = 1
            , @CustomID  = '0052'
            , @nChkArticle  = 0
            , @ArticleID  = '', @TaxClss = '9'

*/




    SET NOCOUNT ON


    -- 기초잔액일자와 조회시작일자 비교 재 설정
    DECLARE @BaseDate   CHAR(8)


    SET  @BaseDate = convert( varchar(8),  dateadd( day,  -1, CONVERT( datetime, @sDate , 12) ) , 112)

--    SELECT @BaseDate = ISNULL(MAX(BasisDate), '')
--      FROM sb_Stock

--    IF @sDate <= @BaseDate 
--       BEGIN
--            SELECT CONVERT( CHAR(8), DATEADD(D, 1, CONVERT( DATETIME, @BaseDate ) ), 112 )
--       END


    -- 수불레코드 ( 입고 & 출고 ) 
    SELECT AA.*
      INTO #SUBUL
      FROM (  -- 입고데이터
             SELECT kCustom = DD.KCustom, CC.Article, AA.SubulWidthID, IODate = AA.StuffDate, Cls = '1', AA.Custom, StuffRoll = AA.TotRoll, StuffQty= AA.TotQty
                  , OrderNO = '', OutRoll=0, OutQtyYDS = 0, UnitClss='', OutQty=0, OutRealQty = 0, Remark=AA.Remark
                  , pkey = AA.StuffDate + '-' + AA.StuffClss + '-' + Convert( Varchar(4), AA.StuffSeq )
                  , CASE LEN(RTRIM(AA.Memo))
			WHEN 0 THEN 0
                        ELSE 1
                    END AS Memo		
                  , SubulWidth = ISNULL( ( SELECT StuffWidth
                                             FROM mt_StuffWidth ZZ
                                            WHERE AA.SubulWidthID = ZZ.StuffWidthID), '' )
               FROM [StuffIN] AA, [mt_Article] CC, [mt_Custom] DD
              WHERE AA.ArticleID = CC.ArticleID
                AND AA.CustomID = DD.CustomID
                AND AA.StuffDate BETWEEN @sDate  AND @eDate
                AND ( ( @nChkCustom = 0 )
                   OR ( @nChkCustom = 1 AND AA.CustomID = @CustomID ) )
                AND ( ( @nChkArticle = 0 )
                   OR ( @nChkArticle = 1 AND AA.ArticleID = @ArticleID  AND AA.SubulWidthID = @SubulWidthID ) )
                AND ( ( @TaxClss = '9' )
                   OR ( @TaxClss = '0' AND AA.OrderFlag = @TaxClss ) 
                   OR ( @TaxClss = '1' AND AA.OrderFlag = @TaxClss ) )

              UNION ALL  


             -- 출고데이터
             SELECT kCustom = DD.KCustom, CC.Article, AA.SubulWidthID, IODate = AA.OutDate, Cls= '2', Custom=OutCustom, StuffRoll=0, StuffQty=0
                  , OrderNO = CASE AA.OutClss WHEN '2' THEN '제직불량' 
						WHEN '3' THEN '가공불량'
						WHEN '4' THEN 'Sample'
						WHEN '5' THEN '정산분'
						ELSE BB.OrderNO END 
		  , AA.OutRoll
                  , CASE BB.UnitClss
                          WHEN  0  THEN AA.OutQty 
                          ELSE          CONVERT( INT, ROUND(AA.OutQty / 0.9144, 0 ) )
                    END AS OutQtyYDS
                  , CASE BB.UnitClss
                          WHEN  '0'  THEN ' Y'
                          WHEN  '1'  THEN ' M'
                          ELSE       ' K'
                    END AS UnitClss
                  , AA.OutQty
                  , AA.OutRealQty, Remark= AA.Remark
                  , pkey = AA.OrderID + '-' + Convert( Varchar(4), AA.OutSeq)
                  , CASE LEN(RTRIM(AA.Memo))
			WHEN 0 THEN 0
                        ELSE 1
                    END AS Memo				 
                  , SubulWidth = ISNULL( ( SELECT StuffWidth
                                             FROM mt_StuffWidth ZZ
                                            WHERE AA.SubulWidthID = ZZ.StuffWidthID ), '' )
               FROM [OutWare] AA, [Order] BB, [mt_Article] CC, [mt_Custom] DD
              WHERE AA.OrderID = BB.OrderID
                AND BB.ArticleID = CC.ArticleID 
                AND BB.CustomID = DD.CustomID
                AND AA.OutDate BETWEEN @sDate  AND @eDate
                AND ( ( @nChkCustom = 0 )
                   OR ( @nChkCustom = 1 AND BB.CustomID = @CustomID ) )
                AND ( ( @nChkArticle = 0 )
                   OR ( @nChkArticle = 1 AND BB.ArticleID = @ArticleID  AND AA.SubulWidthID = @SubulWidthID ) )
                AND ( ( @TaxClss = '9' )
                   OR ( @TaxClss = '0' AND BB.OrderFlag = @TaxClss ) 
                   OR ( @TaxClss = '1' AND BB.OrderFlag = @TaxClss ) )
          ) AA



    -- 전기이월 레코드 
    SELECT DISTINCT kCustom=BB.KCustom, CC.Article, AA.SubulWidthID, IODate = ' ', Cls = '0', Custom=' ', StuffRoll=0, StuffQty = ISNULL( AA.StockQty +  AA.StuffINQty -  AA.OutQty, 0 )
         , OrderNO=' ', OutRoll=0, OutQty=0, OutRealQty=0, Remark=' ', pkey=' ', Memo=0
         , SubulWidth = ISNULL( ( SELECT StuffWidth
                                    FROM mt_StuffWidth ZZ
                                   WHERE AA.SubulWidthID = ZZ.StuffWidthID ), '' )

      INTO #Stock 
      FROM DBO.fn_sbStock_sStock(@BaseDate, @nChkCustom, @CustomID, @nChkArticle, @ArticleID, @SubulWidthID, @TaxClss )  AA
         , [mt_Custom] BB, [mt_Article] CC   --, #SUBUL DD
     WHERE AA.CustomID = BB.CustomID
       AND AA.ArticleID = CC.ArticleID  
--       AND BB.kCustom = DD.kCustom
--       AND CC.Article = DD.Article

   -- 전기이월  + 수불 레코드 
   SELECT *
     INTO #TEMP
     FROM ( SELECT kCustom, Article, SubulWidth, IODate, Cls, Custom, StuffRoll, StuffQty
                 , OrderNO, OutRoll, OutQty, OutRealQty, Remark, Pkey, Memo, UnitClss='', OutQtyYDS=0
             FROM #Stock 
            WHERE StuffQty <> 0

             UNION ALL

            SELECT kCustom, Article, SubulWidth, IODate, Cls, Custom, StuffRoll, StuffQty
                 , OrderNO, OutRoll, OutQty, OutRealQty, Remark, Pkey, Memo, UnitClss, OutQtyYDS
              FROM #SUBUL
          ) AS CC


   -- SELECT 레코드
   SELECT kCustom, Article, SubulWidth, IODate, Cls, Custom, StuffRoll, StuffQty
        , OrderNO, OutRoll, OutQty, OutRealQty, Remark, Pkey, Memo, UnitClss
     FROM #TEMP

    UNION ALL

   SELECT kCustom, Article, SubulWidth, IODate= '99999999', Cls='3', Custom='', StuffRoll = SUM(StuffRoll), StuffQty = SUM(StuffQty)
        , OrderNO='', OutRoll=SUM(OutRoll), OutQty=SUM(OutQtyYDS), OutRealQty=SUM(OutRealQty), Remark='', pkey='', Memo=0, UnitClss='  '
     FROM #TEMP
 GROUP BY kCustom, Article, SubulWidth 
 ORDER BY kCustom, Article, SubulWidth, IODate








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





--  exec xp_StuffIN_sStuffINByOrder  '2004112717'



ALTER           PROCEDURE  xp_StuffIN_sStuffINByOrder
	@OrderID    char(10)
AS
  SELECT AA.OrderID, AcptDate, CustomID, AA.OrderNO, AA.ArticleID
       , Article = ISNULL((SELECT Article
                             FROM mt_Article BB
                            WHERE AA.ArticleID = BB.ArticleID), '' )
       , WorkName = ISNULL((SELECT WorkName
                              FROM mt_work BB
                             WHERE AA.WorkID = BB.WorkID), '' )
       , LossRate, ChunkRate, ColorCnt, OrderQty
       , StuffWidth = ISNULL((SELECT StuffWidth
                                FROM mt_stuffwidth BB
                               WHERE AA.StuffWidth = BB.StuffWidthID), '' )
       , LossRate, ChunkRate, ColorCnt, OrderQty
       , InRoll = ISNULL(CC.InputRoll,0), InQty = ISNULL(CC.InputQty, 0)
       , CASE AA.UnitClss
              WHEN 0 THEN 'Y'
              WHEN 1 THEN 'M'
              ELSE ''
         END AS UnitClss
       , kCustom = ISNULL(( SELECT kCustom
                              FROM mt_custom  DD
                             WHERE DD.CustomID = AA.CustomID ), '' )
       , Width = ISNULL(( SELECT StuffWidth
                              FROM mt_stuffwidth  DD
                             WHERE DD.StuffWidthID = AA.workwidth ), '' )
       , NeedQty = CASE AA.UnitClss
                       WHEN '0'  THEN AA.OrderQty * ( 1 + LossRate/100 + ChunkRate / 100 )
                       WHEN '1'  THEN AA.OrderQty * 1.0936 * ( 1 + LossRate/100 + ChunkRate / 100 )
                   END 
       , NonInQty = CASE AA.UnitClss
                       WHEN '0'  THEN ( AA.OrderQty * ( 1 + LossRate/100 + ChunkRate / 100 ) )  - ISNULL(CC.InputQty, 0)
                       WHEN '1'  THEN ( AA.OrderQty * 1.0936 * ( 1 + LossRate/100 + ChunkRate / 100 ) ) - ISNULL(CC.InputQty, 0)
                   END 
       , AA.OrderFlag
       , CC.TaxClss
       , AA.SubulWidthID, BB.StuffWidth
--( AA.OrderQty * ( 1 + LossRate/100 + ChunkRate / 100 ) ) - ISNULL(CC.InputQty, 0)
    FROM [Order] AA   LEFT OUTER JOIN 
        ( SELECT BB.OrderID, InputRoll= ISNULL( SUM(TotRoll),0) , InputQty = isnull(sum(TotQty),0), TaxClss = MAX(OrderFlag)
             FROM StuffIn BB
            WHERE BB.OrderID = @OrderID
         GROUP BY BB.OrderID
         ) AS  CC ON  AA.OrderID = CC.OrderID
       , [mt_StuffWidth] BB
   WHERE AA.OrderID = @OrderID
     AND AA.SubulWidthID = BB.StuffWidthID













GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/********************************************************
 * Date : 2002-10-12 (SAT)
 *
 * Description: 생지 입고 추가 - 최현숙
 *		(CStuffIN -> frmStuffIN)
 ********************************************************/
ALTER                PROCEDURE xp_StuffIN_iuStuffIN
    @nAfftecRows   smallint     OUTPUT
  , @JobFlag       char(1)
  , @StuffDate     char(8) = ''
  , @StuffClss     char(1) = ''
  , @StuffSeq      smallint = 0  OUTPUT
  , @CustomID      char(4) = ''
  , @Custom        nvarchar(40) = ''
  , @UnitClss      char(1) = '' 
  , @TotRoll       smallint = 0
  , @TotQty        integer = 0
  , @Remark        nvarchar(200) = ''
  , @ThreadName    nvarchar(30) = ''
  , @OrderId       char(10) = ''
  , @ArticleID     char(4) = ''
  , @WorkID        char(4) = ''
  , @OrderNO       nvarchar(24) = ''
  , @AddClss       char(1) = ''
  , @sOrderFlag    CHAR(1) = '0'
  , @UpdDateFlag   CHAR(1) = '0'   -- 입고일자 수정 0: 비수정,   1: 수정
  , @OldStuffDate  CHAR(8) = ''    -- 
  , @OldStuffClss  CHAR(1) = ''
  , @OldStuffSeq   smallint = 0
  , @SubulWidthID  CHAR(2) = ''
AS
       DECLARE @TotAssign  INTEGER

       SET @TotAssign = 0

       IF RTRIM( @OrderNO ) <> '' 
           SET @TotAssign = @TotQty 

       
       IF @UpdDateFlag = '1'      -- 입고일자 수정
           BEGIN
               SET @StuffSeq = dbo.fn_NewStuffSeq(@StuffDate, @StuffClss )

               INSERT [StuffIN] (StuffDate, StuffClss, StuffSeq, CustomID, Custom, UnitClss, TotRoll, TotQty, Remark, CurrDate
                               , ThreadName, OrderId, ArticleID, WorkID, OrderNO, ADDClss, OrderFlag, AssignQty, SubulWidthID ) 
               VALUES ( @StuffDate, @StuffClss, @StuffSeq, @CustomID, @Custom, @UnitClss, @TotRoll, @TotQty, @Remark, GETDATE() 
                      , @ThreadName, @OrderId, @ArticleID, @WorkID, @OrderNO, @ADDClss, @sOrderFlag, @TotAssign, @SubulWidthID)

               UPDATE [StuffAssign]
                  SET StuffDate = @StuffDate
                    , StuffClss = @StuffClss
                    , StuffSeq = @StuffSeq
                WHERE StuffDate = @OldStuffDate
                  AND StuffClss = @OldStuffClss
                  AND StuffSeq = @OldStuffSeq

               UPDATE [StuffINSub]
                  SET StuffDate = @StuffDate
                    , StuffClss = @StuffClss
                    , StuffSeq = @StuffSeq
                WHERE StuffDate = @OldStuffDate
                  AND StuffClss = @OldStuffClss
                  AND StuffSeq = @OldStuffSeq
                
               UPDATE [StuffINReturn]
                  SET StuffDate = @StuffDate
                    , StuffClss = @StuffClss
                    , StuffSeq = @StuffSeq
                WHERE StuffDate = @OldStuffDate
                  AND StuffClss = @OldStuffClss
                  AND StuffSeq = @OldStuffSeq

               DELETE [StuffIN]
                WHERE StuffDate = @OldStuffDate
                  AND StuffClss = @OldStuffClss
                  AND StuffSeq = @OldStuffSeq
                 

           END


       ELSE
           -- 순수, 입력/수정일 경우
           BEGIN

                 IF @JobFlag = 'I'
                     BEGIN
                         SET @StuffSeq = dbo.fn_NewStuffSeq(@StuffDate, @StuffClss )

                         INSERT [StuffIN] (StuffDate, StuffClss, StuffSeq, CustomID, Custom, UnitClss, TotRoll, TotQty, Remark, CurrDate
                                         , ThreadName, OrderId, ArticleID, WorkID, OrderNO, ADDClss, OrderFlag, AssignQty, SubulWidthID ) 
                         VALUES ( @StuffDate, @StuffClss, @StuffSeq, @CustomID, @Custom, @UnitClss, @TotRoll, @TotQty, @Remark, GETDATE() 
                                , @ThreadName, @OrderId, @ArticleID, @WorkID, @OrderNO, @ADDClss, @sOrderFlag, @TotAssign, @SubulWidthID)

                     END
                 ELSE
                     BEGIN
                         UPDATE [StuffIN]
                            SET CustomID  = @CustomID
                              , Custom = @Custom
                              , UnitClss = @UnitClss
                              , TotRoll = @TotRoll
                              , TotQty = @TotQty
                              , Remark = @Remark
                              , CurrDate = GETDATE()
                              , ThreadName = @ThreadName
                              , OrderId = @OrderId
                              , ArticleID = @ArticleID
                              , WorkID = @WorkID
                              , OrderNO = @OrderNO
                              , ADDClss = @ADDClss
                              , OrderFlag = @sOrderFlag
                              , AssignQty = @TotAssign
                              , SubulWidthID = @SubulWidthID
                          WHERE StuffDate = @StuffDate
                            AND StuffClss = @StuffClss
                            AND StuffSeq = @StuffSeq
 
                         DELETE [StuffINSub]
                          WHERE StuffDate = @StuffDate
                            AND StuffClss = @StuffClss
                            AND StuffSeq = @StuffSeq

                         DELETE [StuffINReturn]
                          WHERE StuffDate = @StuffDate
                            AND StuffClss = @StuffClss
                            AND StuffSeq = @StuffSeq
                          
 
                    END
                 SET @nAfftecRows = @@ROWCOUNT
          END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO






ALTER      PROCEDURE xp_iuStuffAssign
    @JobFlag      CHAR(1)
  , @StuffDate    CHAR(8)
  , @StuffClss    CHAR(1)
  , @StuffSeq     INT
  , @OrderID      CHAR(10)
  , @AssignSeq    INT
  , @Qty          INT
  , @ROLL         INT
  , @AssignDate   CHAR(8)
AS
    DECLARE @SubulWidthID CHAR(2)

    SELECT @SubulWidthID = WorkWidth 
      FROM [Order]
     WHERE OrderID = @OrderID

    IF @JobFlag = 'I'
    	
        INSERT INTO StuffAssign( StuffDate, StuffClss, StuffSeq, OrderID, AssignSeq, Qty, Roll, AssignDate, SetDate, SubulWidthID )
        VALUES ( @StuffDate, @StuffClss, @StuffSeq, @OrderID
                , ISNULL( dbo.fn_NewStuffAssignSeq( @StuffDate, @StuffClss, @StuffSeq, @OrderID ), 1 ) 
                , @Qty, @Roll, @AssignDate, getdate(), @SubulWidthID )
    ELSE
        UPDATE StuffAssign
           SET Qty = @Qty
             , Roll = @Roll
             , AssignDate = @AssignDate
             , SetDate =  getdate() 
             , SubulWidthID = @SubulWidthID
         WHERE StuffDate = @StuffDate
           AND StuffClss =  @StuffClss
           AND StuffSeq =  @StuffSeq
           AND OrderID =  @OrderID
           AND AssignSeq = @AssignSeq


    UPDATE StuffIN
       SET SubulWidthID = @SubulWidthID
     WHERE StuffDate = @StuffDate
       AND StuffClss =  @StuffClss
       AND StuffSeq =  @StuffSeq
        

    IF @@ROWCOUNT = 1 
         EXEC xp_StuffIN_uAssignQty   @StuffDate, @StuffClss, @StuffSeq







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







ALTER      PROCEDURE  xp_StuffIN_sStuffINONE
	@StuffDate  CHAR(8)
      , @StuffClss  CHAR(1)
      , @StuffSeq   Smallint 
AS
	SELECT StuffDate
	     , StuffClss
	     , StuffSeq
	     , CustomID
             , kCustom = ISNULL((SELECT kCustom
                               FROM mt_Custom  BB
                              WHERE AA.CustomID = BB.CustomID), '' )
                               
	     , Custom                                   
	     , OrderID     
             , OrderNO
	     , ArticleID 
	     , Article = ISNULL((SELECT Article
	                           FROM mt_Article BB
	                          WHERE AA.ArticleID = BB.ArticleID), '' )
	     , WorkID 
             , WorkName = ISNULL((SELECT WorkName
                                    FROM mt_work BB
                                   WHERE AA.WorkID = BB.WorkID), '' )
	     , UnitClss 
             , CASE AA.UnitClss
                    WHEN 0 THEN 'YDS'
                    WHEN 1 THEN 'MTS'
                    ELSE ''
               END AS UnitName
	     , TotRoll, TotQty, Remark, ThreadName, AddClss, OrderFlag
             , AssignQty = ISNULL( ( SELECT SUM(QTY)
                                       FROM StuffAssign ZZ
                                      WHERE ZZ.StuffDate = AA.StuffDate
                                        AND ZZ.StuffClss = AA.StuffClss
                                        AND ZZ.StuffSeq = AA.StuffSeq ), 0 )
             , SubulWidthID                                           
	  FROM StuffIN AA
	 WHERE StuffDate = @StuffDate
	   AND StuffClss = @StuffClss
	   AND StuffSeq = @StuffSeq








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

